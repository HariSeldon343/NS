#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoraggio Atti Calabria
Scarica e filtra gli atti da:
  - 5 portali ASP Calabria (AlboOnline)
  - Regione Calabria (provvedimenti)
  - TAR Catanzaro (giustizia-amministrativa)
Genera un file XLSX con un foglio per portale.
"""

import re
import sys
import time
import warnings
from datetime import date, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------

TODAY = date.today()
DATE_FROM_2D = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM_3D = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")
DATE_TO      = TODAY.strftime("%Y-%m-%d")

KEYWORDS_PLAIN = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]

# Costruisce il pattern unico: Life con word-boundary, resto con semplice contains
def _build_pattern(keywords: list[str]) -> re.Pattern:
    parts = []
    for kw in keywords:
        if kw.upper() == "LIFE":
            parts.append(r"\bLIFE\b")
        else:
            parts.append(re.escape(kw))
    return re.compile("|".join(parts), re.IGNORECASE)

KEYWORD_PATTERN = _build_pattern(KEYWORDS_PLAIN + ["LIFE"])

PORTALS_ASP = [
    ("ASP Cosenza",         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Catanzaro",       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Crotone",         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Reggio Calabria", "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Vibo Valentia",   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
]

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
})

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def matches(text: str) -> bool:
    return bool(KEYWORD_PATTERN.search(text))

def short_desc(text: str, n: int = 150) -> str:
    t = text.strip()
    return t[:n] + "…" if len(t) > n else t

def get_page(url: str, params: dict | None = None, data: dict | None = None,
             method: str = "GET", timeout: int = 30, retries: int = 3) -> requests.Response | None:
    for attempt in range(retries):
        try:
            if method.upper() == "POST":
                r = SESSION.post(url, data=data, params=params,
                                 timeout=timeout, verify=False)
            else:
                r = SESSION.get(url, params=params, timeout=timeout, verify=False)
            r.raise_for_status()
            return r
        except requests.exceptions.Timeout:
            print(f"    Timeout (tentativo {attempt+1}/{retries})")
            time.sleep(2 ** attempt)
        except requests.exceptions.HTTPError as e:
            print(f"    HTTP Error: {e}")
            return None
        except requests.exceptions.ConnectionError as e:
            print(f"    Connessione fallita: {e}")
            return None
    return None

# ---------------------------------------------------------------------------
# Scraper ASP Calabria (AlboOnline)
# ---------------------------------------------------------------------------

def _asp_parse_page(html: str, base_url: str) -> list[dict]:
    """Estrae righe dalla tabella risultati dell'AlboOnline."""
    soup = BeautifulSoup(html, "lxml")
    rows = []

    # La tabella principale ha class 'table' o simili
    table = soup.find("table")
    if not table:
        return rows

    header_cells = table.find("tr")
    if not header_cells:
        return rows
    headers = [th.get_text(strip=True).lower() for th in header_cells.find_all(["th", "td"])]

    def col_idx(names):
        for n in names:
            for i, h in enumerate(headers):
                if n in h:
                    return i
        return None

    idx_data  = col_idx(["data pubbl", "data di pubbl", "data"])
    idx_obj   = col_idx(["oggetto"])
    idx_link  = col_idx(["dettaglio", "allegato", "documento", "azione"])

    for tr in table.find_all("tr")[1:]:
        cells = tr.find_all("td")
        if not cells:
            continue
        n = len(cells)

        data_val = cells[idx_data].get_text(strip=True) if idx_data is not None and idx_data < n else ""
        obj_val  = cells[idx_obj].get_text(strip=True)  if idx_obj  is not None and idx_obj  < n else ""

        # Cerca link in tutta la riga
        link = ""
        a_tags = tr.find_all("a", href=True)
        if a_tags:
            href = a_tags[-1]["href"]
            link = href if href.startswith("http") else base_url.rstrip("/") + "/" + href.lstrip("/")

        rows.append({"data": data_val, "oggetto": obj_val, "link": link})
    return rows

def _asp_get_total_pages(soup: BeautifulSoup) -> int:
    """Stima numero pagine dall'elemento di paginazione."""
    # Cerca pattern comuni: "Pagina X di Y", select con option, ul.pagination
    # Pattern 1: testo "Pagina X di Y"
    for tag in soup.find_all(string=re.compile(r"pagina\s+\d+\s+di\s+\d+", re.I)):
        m = re.search(r"di\s+(\d+)", tag, re.I)
        if m:
            return int(m.group(1))
    # Pattern 2: ul.pagination
    pag = soup.find("ul", class_=re.compile(r"pagination", re.I))
    if pag:
        nums = [int(a.get_text(strip=True)) for a in pag.find_all("a")
                if a.get_text(strip=True).isdigit()]
        if nums:
            return max(nums)
    # Pattern 3: select di pagina
    sel = soup.find("select", id=re.compile(r"page", re.I))
    if sel:
        opts = [o.get("value", "") for o in sel.find_all("option") if o.get("value","").isdigit()]
        if opts:
            return max(int(o) for o in opts)
    return 1

def scrape_asp(name: str, base_url: str) -> list[dict]:
    """Scarica tutti gli atti AlboOnline per gli ultimi 2 giorni."""
    print(f"\n[{name}] Inizio scraping...")
    results = []

    # Prima richiesta per ottenere viewstate e struttura form
    r0 = get_page(base_url)
    if r0 is None:
        print(f"  ERRORE: portale non raggiungibile")
        return [{"_error": f"Portale {name} non raggiungibile (timeout/connessione)"}]

    soup0 = BeautifulSoup(r0.text, "lxml")

    # Raccoglie campi hidden necessari (ASP.NET viewstate, ecc.)
    form = soup0.find("form")
    post_data = {}
    if form:
        for inp in form.find_all("input"):
            if inp.get("type") in ("hidden", None) and inp.get("name"):
                post_data[inp["name"]] = inp.get("value", "")

    # Imposta la data di pubblicazione "dal"
    # Cerca il campo con id="dataPubblicazioneDal" o simili
    date_field_names = []
    if form:
        for inp in form.find_all("input"):
            fid = inp.get("id", "").lower()
            fname = inp.get("name", "").lower()
            if "datapubbl" in fid or "datapubbl" in fname or "dal" in fid:
                date_field_names.append(inp.get("name") or inp.get("id"))

    # Parametri di ricerca
    search_params = dict(post_data)
    for fn in date_field_names:
        search_params[fn] = DATE_FROM_2D

    # Prova metodo POST (AlboOnline tipicamente usa POST)
    r1 = get_page(base_url, data=search_params, method="POST")
    if r1 is None:
        # Riprova GET
        r1 = get_page(base_url, params={
            "dataPubblicazioneDal": DATE_FROM_2D,
            "cerca": "1",
        })
    if r1 is None:
        print(f"  ERRORE: ricerca non riuscita")
        return [{"_error": f"Portale {name}: ricerca non riuscita"}]

    soup1 = BeautifulSoup(r1.text, "lxml")
    total_pages = _asp_get_total_pages(soup1)
    print(f"  Pagine trovate: {total_pages}")

    for page_num in range(1, total_pages + 1):
        print(f"  Pagina {page_num}/{total_pages}...")
        if page_num == 1:
            html = r1.text
        else:
            # Paginazione: prova param page/pagina/p
            page_data = dict(search_params)
            page_data["pagina"] = str(page_num)
            page_data["page"]   = str(page_num)
            rp = get_page(base_url, data=page_data, method="POST")
            if rp is None:
                rp = get_page(base_url, params={
                    "dataPubblicazioneDal": DATE_FROM_2D,
                    "pagina": str(page_num),
                })
            if rp is None:
                break
            html = rp.text

        page_rows = _asp_parse_page(html, base_url)
        if not page_rows:
            break

        for row in page_rows:
            if matches(row.get("oggetto", "")):
                results.append(row)

        time.sleep(0.5)

    print(f"  Match trovati: {len(results)}")
    return results

# ---------------------------------------------------------------------------
# Scraper Regione Calabria
# ---------------------------------------------------------------------------

RC_BASE = "https://www.regione.calabria.it"
RC_URL  = f"{RC_BASE}/provvedimenti-della-regione/"

def scrape_regione_calabria() -> list[dict]:
    print("\n[Regione Calabria] Inizio scraping...")
    results = []
    page = 1

    while True:
        print(f"  Pagina {page}...")
        params = {
            "filter_date_from": DATE_FROM_2D,
            "filter_date_to":   DATE_TO,
            "filter_active":    "true",
            "paged":            str(page),
        }
        r = get_page(RC_URL, params=params)
        if r is None:
            print("  ERRORE: pagina non raggiungibile")
            break

        soup = BeautifulSoup(r.text, "lxml")
        table = soup.find("table")
        if not table:
            print("  Nessuna tabella trovata")
            break

        data_rows = table.find_all("tr")[1:]
        if not data_rows:
            break

        new_found = 0
        for tr in data_rows:
            cells = tr.find_all("td")
            if len(cells) < 4:
                continue

            # Colonne: Tipologia | Data | Dipartimento | Oggetto | →
            data_val = cells[1].get_text(strip=True)
            # Col 2 = Dipartimento, Col 3 = Oggetto
            obj_val  = cells[3].get_text(strip=True)

            # Link
            a = tr.find("a", href=True)
            link = a["href"] if a else ""
            if link and not link.startswith("http"):
                link = RC_BASE + link

            if matches(obj_val):
                results.append({"data": data_val, "oggetto": obj_val, "link": link})
                new_found += 1

        print(f"    Righe: {len(data_rows)}, match: {new_found}")

        # Controlla se c'è pagina successiva
        pag = soup.find("ul", class_="page-numbers")
        has_next = False
        if pag:
            next_link = pag.find("a", class_="next")
            has_next = next_link is not None
        if not has_next:
            break

        page += 1
        time.sleep(0.3)

    print(f"  Totale match: {len(results)}")
    return results

# ---------------------------------------------------------------------------
# Scraper TAR Catanzaro
# ---------------------------------------------------------------------------

TAR_BASE = "https://www.giustizia-amministrativa.it"
TAR_URL  = f"{TAR_BASE}/provvedimenti-tar-catanzaro"

def scrape_tar_catanzaro() -> list[dict]:
    print("\n[TAR Catanzaro] Inizio scraping...")
    results = []
    page = 1

    while True:
        print(f"  Pagina {page}...")
        params = {
            "publishDateFrom": DATE_FROM_3D,
            "publishDateTo":   DATE_TO,
            "page":            str(page),
        }
        r = get_page(TAR_URL, params=params)
        if r is None:
            print("  ERRORE: portale non raggiungibile")
            return [{"_error": "Portale TAR Catanzaro non raggiungibile"}]

        soup = BeautifulSoup(r.text, "lxml")

        # Cerca tabella risultati
        table = soup.find("table")
        if not table:
            # Cerca lista alternativa
            items = soup.find_all("div", class_=re.compile(r"result|item|row", re.I))
            if not items:
                print("  Nessun risultato trovato")
                break

        if table:
            data_rows = table.find_all("tr")[1:]
            if not data_rows:
                break

            new_found = 0
            for tr in data_rows:
                cells = tr.find_all("td")
                if not cells:
                    continue

                # Cerca colonna "Parte"
                parte_val = ""
                data_val  = ""
                link      = ""

                for i, c in enumerate(cells):
                    txt = c.get_text(strip=True)
                    # Colonna data
                    if re.match(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{4}", txt):
                        data_val = txt
                    # Colonna Parte (tipicamente contiene il nome della parte)
                    elif len(txt) > 3 and not txt.isdigit():
                        if not parte_val:
                            parte_val = txt

                a = tr.find("a", href=True)
                if a:
                    href = a["href"]
                    link = href if href.startswith("http") else TAR_BASE + href

                if matches(parte_val):
                    results.append({
                        "data": data_val,
                        "oggetto": parte_val,
                        "link": link,
                    })
                    new_found += 1

            print(f"    Righe: {len(data_rows)}, match: {new_found}")

        # Controlla paginazione
        pag = soup.find("ul", class_=re.compile(r"pagination", re.I))
        has_next = False
        if pag:
            next_a = pag.find("a", string=re.compile(r"next|success|>|»", re.I))
            if not next_a:
                next_li = pag.find("li", class_=re.compile(r"next", re.I))
                if next_li and next_li.find("a"):
                    next_a = next_li.find("a")
            has_next = next_a is not None and "disabled" not in (next_li.get("class", []) if pag.find("li", class_=re.compile(r"next", re.I)) else [])
        if not has_next:
            break
        page += 1
        time.sleep(0.3)

    print(f"  Totale match: {len(results)}")
    return results

# ---------------------------------------------------------------------------
# Generazione Excel
# ---------------------------------------------------------------------------

HEADER_FILL  = PatternFill("solid", fgColor="1F5C99")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
LINK_FONT    = Font(color="0563C1", underline="single", size=10)
ERROR_FONT   = Font(italic=True, color="CC0000", size=10)
ALT_FILL     = PatternFill("solid", fgColor="EBF3FB")
EMPTY_FILL   = PatternFill("solid", fgColor="F5F5F5")

def _setup_sheet(ws, portal_name: str):
    ws.title = portal_name[:31]  # max 31 chars per Excel
    ws.append(["Data", "Oggetto", "Descrizione breve (150 car.)", "Link all'atto"])
    header_row = ws[1]
    for cell in header_row:
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 22

def _set_col_widths(ws):
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 20

def _write_results(ws, rows: list[dict]):
    if not rows:
        cell = ws.cell(row=2, column=1, value="Nessun risultato")
        cell.font = Font(italic=True, color="888888", size=10)
        return

    # Controlla errore di connessione
    if len(rows) == 1 and "_error" in rows[0]:
        cell = ws.cell(row=2, column=1, value=rows[0]["_error"])
        cell.font = ERROR_FONT
        ws.merge_cells("A2:D2")
        return

    for i, row in enumerate(rows, start=2):
        fill = ALT_FILL if i % 2 == 0 else None

        data_cell = ws.cell(row=i, column=1, value=row.get("data", ""))
        data_cell.alignment = Alignment(horizontal="center", vertical="top")
        if fill: data_cell.fill = fill

        obj_cell = ws.cell(row=i, column=2, value=row.get("oggetto", ""))
        obj_cell.alignment = Alignment(wrap_text=True, vertical="top")
        if fill: obj_cell.fill = fill

        desc_cell = ws.cell(row=i, column=3, value=short_desc(row.get("oggetto", "")))
        desc_cell.alignment = Alignment(wrap_text=True, vertical="top")
        if fill: desc_cell.fill = fill

        link_val = row.get("link", "")
        link_cell = ws.cell(row=i, column=4, value="Apri atto" if link_val else "—")
        if link_val:
            link_cell.hyperlink = link_val
            link_cell.font = LINK_FONT
        link_cell.alignment = Alignment(horizontal="center", vertical="top")
        if fill: link_cell.fill = fill

def generate_excel(all_data: dict[str, list[dict]], output_path: Path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuove foglio default

    for portal_name, rows in all_data.items():
        ws = wb.create_sheet()
        _setup_sheet(ws, portal_name)
        _write_results(ws, rows)
        _set_col_widths(ws)

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("MONITORAGGIO ATTI CALABRIA")
    print(f"Data odierna: {TODAY}")
    print(f"Intervallo ASP/RC: {DATE_FROM_2D} → {DATE_TO}")
    print(f"Intervallo TAR:    {DATE_FROM_3D} → {DATE_TO}")
    print("=" * 60)

    all_data: dict[str, list[dict]] = {}

    # --- ASP portali ---
    for portal_name, portal_url in PORTALS_ASP:
        all_data[portal_name] = scrape_asp(portal_name, portal_url)

    # --- Regione Calabria ---
    all_data["Regione Calabria"] = scrape_regione_calabria()

    # --- TAR Catanzaro ---
    all_data["TAR Catanzaro"] = scrape_tar_catanzaro()

    # --- Excel output ---
    file_name = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"

    # Percorso output: usa argomento o cartella corrente
    if len(sys.argv) > 1:
        output_dir = Path(sys.argv[1])
    else:
        output_dir = Path(".")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / file_name

    generate_excel(all_data, output_path)

    # Riepilogo
    print("\n--- RIEPILOGO ---")
    for name, rows in all_data.items():
        if rows and "_error" in rows[0]:
            status = f"ERRORE: {rows[0]['_error']}"
        else:
            status = f"{len(rows)} match"
        print(f"  {name}: {status}")

if __name__ == "__main__":
    main()
