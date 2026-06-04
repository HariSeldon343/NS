#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraping multi-portale con filtering per keyword sanitarie.
"""

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import time
import os
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─────────────────────────────────────────────────────────
# CONFIGURAZIONE
# ─────────────────────────────────────────────────────────

TODAY = datetime(2026, 6, 4)
DATE_FROM_2 = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")   # 2026-06-02
DATE_FROM_3 = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")   # 2026-06-01
DATE_TO     = TODAY.strftime("%Y-%m-%d")                         # 2026-06-04

OUTPUT_FILENAME = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
OUTPUT_PATH = os.path.join("/home/user/NS", OUTPUT_FILENAME)

KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo",
    "Autorizzazione all'esercizio", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni",
    "Budget", "Casa Giardino", "Centro San Giuseppe",
    "Centro salute e benessere", "Fabbisogni LEA",
    "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale",
    "Riabilitazione estensiva", "Riconversione prestazioni",
    "Rinnovo accreditamento", "San Teodoro", "Savelli Hospital",
    "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]

LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

SHEET_ORDER = [
    "ASP Cosenza",
    "ASP Catanzaro",
    "ASP Crotone",
    "ASP Reggio Calabria",
    "ASP Vibo Valentia",
    "Regione Calabria",
    "TAR Catanzaro",
]

HEADERS_MAP = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9",
    "Accept-Encoding": "gzip, deflate",
}


# ─────────────────────────────────────────────────────────
# KEYWORD MATCHING
# ─────────────────────────────────────────────────────────

def matches(text: str) -> bool:
    if not text:
        return False
    tu = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in tu:
            return True
    return bool(LIFE_RE.search(text))


# ─────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────

def abs_url(href: str, base: str) -> str:
    if not href:
        return ""
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        from urllib.parse import urlparse
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{href}"
    return base.rstrip("/") + "/" + href


def safe_get(session, url, params=None, data=None, method="GET", timeout=30):
    try:
        if method == "POST":
            r = session.post(url, data=data, timeout=timeout, verify=False,
                             allow_redirects=True)
        else:
            r = session.get(url, params=params, timeout=timeout, verify=False,
                            allow_redirects=True)
        r.raise_for_status()
        return r
    except requests.exceptions.Timeout:
        return None
    except requests.exceptions.HTTPError as e:
        print(f"    HTTP error: {e}")
        return None
    except Exception as e:
        print(f"    Connection error: {type(e).__name__}: {str(e)[:80]}")
        return None


# ─────────────────────────────────────────────────────────
# ASP PORTALS  (SISR – sisr.regione.calabria.it)
# ─────────────────────────────────────────────────────────

ASP_PORTALS = [
    {
        "name": "ASP Cosenza",
        "base": "https://online-aspco.sisr.regione.calabria.it",
    },
    {
        "name": "ASP Catanzaro",
        "base": "https://online-aspcz.sisr.regione.calabria.it",
    },
    {
        "name": "ASP Crotone",
        "base": "https://online-aspkr.sisr.regione.calabria.it",
    },
    {
        "name": "ASP Reggio Calabria",
        "base": "https://online-asprc.sisr.regione.calabria.it",
    },
    {
        "name": "ASP Vibo Valentia",
        "base": "https://online-aspvv.sisr.regione.calabria.it",
    },
]


def parse_asp_table(soup, base_url):
    """
    Estrae righe dalla tabella risultati dei portali SISR AlboOnline.
    Colonne tipiche: N.Prot | Data | Tipo | Oggetto | [link]
    """
    records = []
    table = None

    # cerca la tabella risultati (evita tabelle di layout/navigazione)
    for t in soup.find_all("table"):
        ths = t.find_all("th")
        tds_first = t.find("tr") and t.find("tr").find_all("td")
        if ths and len(ths) >= 3:
            table = t
            break
        if tds_first and len(tds_first) >= 3:
            table = t
            break

    if not table:
        return records

    rows = table.find_all("tr")
    # Identifica colonne dall'header
    col_map = {}
    header_row = rows[0] if rows else None
    if header_row:
        for i, th in enumerate(header_row.find_all(["th", "td"])):
        	txt = th.get_text(strip=True).lower()
        	if any(k in txt for k in ("data", "pubblicaz")):
        		col_map["data"] = i
        	elif any(k in txt for k in ("oggetto", "titolo", "descrizione")):
        		col_map["oggetto"] = i
        	elif any(k in txt for k in ("tipo", "natura", "tipologia")):
        		col_map["tipo"] = i

    data_rows = rows[1:] if col_map else rows
    for row in data_rows:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue
        n = len(cells)

        # Cerca link nell'intera riga
        link_href = ""
        for cell in cells:
            a = cell.find("a")
            if a and a.get("href"):
                link_href = abs_url(a["href"], base_url)
                break

        # Data
        data_val = ""
        if "data" in col_map and col_map["data"] < n:
            data_val = cells[col_map["data"]].get_text(strip=True)
        else:
            for cell in cells:
                t = cell.get_text(strip=True)
                if re.match(r'\d{2}[/\-\.]\d{2}[/\-\.]\d{4}', t):
                    data_val = t
                    break

        # Oggetto – cella più lunga
        oggetto_val = ""
        if "oggetto" in col_map and col_map["oggetto"] < n:
            oggetto_val = cells[col_map["oggetto"]].get_text(strip=True)
        else:
            for cell in cells:
                t = cell.get_text(strip=True)
                if len(t) > len(oggetto_val) and not re.match(r'^\d{2}[/\-]\d{2}', t):
                    oggetto_val = t

        if oggetto_val:
            records.append({
                "data": data_val,
                "oggetto": oggetto_val,
                "link": link_href,
            })

    return records


def scrape_asp(portal):
    """Scrapa un portale ASP SISR iterando tutte le pagine."""
    name = portal["name"]
    base = portal["base"]
    search_url = f"{base}/AlboOnline/ricercaAlbo"

    session = requests.Session()
    session.headers.update(HEADERS_MAP)

    print(f"  [{name}] Connessione a {search_url} ...")
    r = safe_get(session, search_url, timeout=25)
    if r is None:
        print(f"  [{name}] PORTALE NON RAGGIUNGIBILE")
        return None   # None = errore di connessione

    soup = BeautifulSoup(r.text, "html.parser")

    # Raccogli campi nascosti (CSRF / ViewState / ecc.)
    form_data = {
        "dataPubblicazioneDal": DATE_FROM_2,
        "dataPubblicazioneAl":  DATE_TO,
    }
    for hidden in soup.find_all("input", {"type": "hidden"}):
        name_h = hidden.get("name")
        val_h  = hidden.get("value", "")
        if name_h:
            form_data[name_h] = val_h

    # Ricava action del form
    form = soup.find("form")
    form_action = search_url
    if form and form.get("action"):
        form_action = abs_url(form["action"], base)

    all_records = []
    page = 1

    while page <= 100:
        if page > 1:
            form_data["page"] = str(page)
            form_data.setdefault("pagina", str(page))

        r2 = safe_get(session, form_action, data=form_data, method="POST", timeout=30)
        if r2 is None:
            break

        soup2 = BeautifulSoup(r2.text, "html.parser")
        page_records = parse_asp_table(soup2, base)

        if not page_records:
            break

        all_records.extend(page_records)
        print(f"  [{name}] Pagina {page}: {len(page_records)} righe")

        # Verifica paginazione
        has_next = False
        for link_text in ["successiv", "avanti", "next", ">>"]:
            nxt = soup2.find("a", string=re.compile(link_text, re.I))
            if nxt:
                has_next = True
                break
        if not has_next:
            # Controlla anche link numerici
            pag_div = soup2.find(class_=re.compile(r'paginat|page.*nav', re.I))
            if pag_div:
                current = pag_div.find(class_=re.compile(r'current|active', re.I))
                if current and current.find_next_sibling("a"):
                    has_next = True
        if not has_next:
            break

        page += 1
        time.sleep(0.4)

    return all_records   # lista vuota = raggiunto ma 0 risultati


# ─────────────────────────────────────────────────────────
# REGIONE CALABRIA
# ─────────────────────────────────────────────────────────

def scrape_regione_calabria():
    base = "https://www.regione.calabria.it"
    session = requests.Session()
    session.headers.update(HEADERS_MAP)

    all_records = []
    page = 1

    while page <= 100:
        params = {
            "filter_date_from": DATE_FROM_2,
            "filter_active":    "true",
            "page":             page,
        }
        r = safe_get(session, f"{base}/provvedimenti-della-regione", params=params, timeout=30)
        if r is None:
            return None

        soup = BeautifulSoup(r.text, "html.parser")
        table = soup.find("table")
        if not table:
            break

        rows = table.find_all("tr")[1:]   # salta header
        if not rows:
            break

        page_records = []
        for row in rows:
            cells = row.find_all(["td", "th"])
            if len(cells) < 5:
                continue

            tipologia = cells[0].get_text(strip=True)
            data_rep  = cells[1].get_text(strip=True)
            # cells[2] = numero
            # cells[3] = dipartimento
            oggetto   = cells[4].get_text(strip=True)

            # Link – può essere nella cella Oggetto o in Dettaglio
            link_href = ""
            for cell in cells:
                a = cell.find("a")
                if a and a.get("href"):
                    link_href = abs_url(a["href"], base)
                    break

            if oggetto:
                page_records.append({
                    "data":      data_rep,
                    "oggetto":   oggetto,
                    "link":      link_href,
                    "tipologia": tipologia,
                })

        if not page_records:
            break

        all_records.extend(page_records)
        print(f"  [Regione Calabria] Pagina {page}: {len(page_records)} righe")

        # Paginazione
        has_next = False
        pag = soup.find(class_=re.compile(r'paginat|pagina|wp-page', re.I))
        if pag:
            nxt = pag.find("a", string=re.compile(r'successiv|next|>>', re.I))
            if nxt:
                has_next = True
        if not has_next:
            # Controlla se ci sono link con numero pagina > corrente
            page_links = soup.find_all("a", href=re.compile(rf'page={page+1}|/page/{page+1}', re.I))
            if page_links:
                has_next = True

        if not has_next:
            break

        page += 1
        time.sleep(0.4)

    return all_records


# ─────────────────────────────────────────────────────────
# TAR CATANZARO
# ─────────────────────────────────────────────────────────

def parse_tar_table(soup, base_url):
    records = []
    table = soup.find("table")
    if not table:
        return records

    rows = table.find_all("tr")
    headers = [th.get_text(strip=True).lower() for th in rows[0].find_all(["th", "td"])] if rows else []

    for row in rows[1:]:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue

        row_dict = {}
        for i, cell in enumerate(cells):
            key = headers[i] if i < len(headers) else f"col{i}"
            row_dict[key] = cell.get_text(strip=True)

        # Link
        link_href = ""
        for cell in cells:
            a = cell.find("a")
            if a and a.get("href"):
                link_href = abs_url(a["href"], base_url)
                break

        # Trova campo "parte" e data
        parte = ""
        data_val = ""
        for k, v in row_dict.items():
            if "parte" in k or "ricorr" in k:
                parte = v
            if "data" in k or "pubbl" in k:
                data_val = v

        oggetto = parte or row_dict.get("oggetto", "") or next(iter(row_dict.values()), "")

        if oggetto:
            records.append({
                "data":    data_val,
                "oggetto": oggetto,
                "parte":   parte,
                "link":    link_href,
            })

    return records


def scrape_tar_catanzaro():
    base = "https://www.giustizia-amministrativa.it"
    session = requests.Session()
    session.headers.update(HEADERS_MAP)

    all_records = []
    page = 1

    while page <= 100:
        params = {
            "publishDateFrom": DATE_FROM_3,
            "publishDateTo":   DATE_TO,
            "p_p_id":          "tarProvvedimentiPortlet",
            "page":            page,
        }
        r = safe_get(session, f"{base}/provvedimenti-tar-catanzaro", params=params, timeout=30)
        if r is None:
            return None

        soup = BeautifulSoup(r.text, "html.parser")
        page_records = parse_tar_table(soup, base)

        if not page_records:
            break

        all_records.extend(page_records)
        print(f"  [TAR Catanzaro] Pagina {page}: {len(page_records)} righe")

        has_next = bool(soup.find("a", string=re.compile(r'successiv|avanti|next|>>', re.I)))
        if not has_next:
            break

        page += 1
        time.sleep(0.4)

    return all_records


# ─────────────────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
LINK_FONT    = Font(color="0563C1", underline="single")
ERROR_FONT   = Font(color="C00000", italic=True)

COL_WIDTHS   = [18, 80, 55, 12]
COL_NAMES    = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]


def write_sheet(ws, records, portal_status: str | None = "ok"):
    """
    portal_status:
      "ok"    → records è una lista (può essere vuota)
      "error" → portale non raggiungibile
    """
    # Header
    for col, (h, w) in enumerate(zip(COL_NAMES, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font  = HEADER_FONT
        cell.fill  = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 20

    if portal_status == "error":
        c = ws.cell(row=2, column=1,
                    value="⚠ Portale non raggiungibile dall'ambiente di esecuzione (HTTP 503 / timeout). "
                          "Accedere manualmente dal browser.")
        c.font = ERROR_FONT
        ws.merge_cells("A2:D2")
        return

    if not records:
        ws.cell(row=2, column=1, value="Nessun risultato nel periodo selezionato")
        ws.merge_cells("A2:D2")
        return

    for idx, rec in enumerate(records, 2):
        oggetto = rec.get("oggetto", "")

        ws.cell(row=idx, column=1, value=rec.get("data", ""))
        ws.cell(row=idx, column=2, value=oggetto).alignment = Alignment(wrap_text=True)
        ws.cell(row=idx, column=3, value=oggetto[:150])

        link = rec.get("link", "")
        lc = ws.cell(row=idx, column=4, value="Apri atto" if link else "")
        if link:
            lc.hyperlink = link
            lc.font = LINK_FONT

    ws.row_dimensions[1].height = 20


# ─────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────

def main():
    print("=" * 65)
    print("MONITORAGGIO ATTI CALABRIA")
    print(f"  Periodo ASP / Regione: {DATE_FROM_2} → {DATE_TO}")
    print(f"  Periodo TAR:           {DATE_FROM_3} → {DATE_TO}")
    print("=" * 65)

    all_data   = {}   # name → list of records
    all_status = {}   # name → "ok" | "error"

    # ── ASP portals ──────────────────────────────────────
    for portal in ASP_PORTALS:
        pname = portal["name"]
        print(f"\n▶ {pname}")
        raw = scrape_asp(portal)
        if raw is None:
            all_status[pname] = "error"
            all_data[pname]   = []
            print(f"  → ERRORE di connessione")
        else:
            filtered = [r for r in raw if matches(r.get("oggetto", ""))]
            all_status[pname] = "ok"
            all_data[pname]   = filtered
            print(f"  → {len(raw)} totali, {len(filtered)} con match keyword")

    # ── Regione Calabria ─────────────────────────────────
    print("\n▶ Regione Calabria")
    raw = scrape_regione_calabria()
    if raw is None:
        all_status["Regione Calabria"] = "error"
        all_data["Regione Calabria"]   = []
        print("  → ERRORE di connessione")
    else:
        filtered = [r for r in raw if matches(r.get("oggetto", ""))]
        all_status["Regione Calabria"] = "ok"
        all_data["Regione Calabria"]   = filtered
        print(f"  → {len(raw)} totali, {len(filtered)} con match keyword")

    # ── TAR Catanzaro ────────────────────────────────────
    print("\n▶ TAR Catanzaro")
    raw = scrape_tar_catanzaro()
    if raw is None:
        all_status["TAR Catanzaro"] = "error"
        all_data["TAR Catanzaro"]   = []
        print("  → ERRORE di connessione")
    else:
        filtered = [
            r for r in raw
            if matches(r.get("parte", "")) or matches(r.get("oggetto", ""))
        ]
        all_status["TAR Catanzaro"] = "ok"
        all_data["TAR Catanzaro"]   = filtered
        print(f"  → {len(raw)} totali, {len(filtered)} con match keyword")

    # ── Excel ────────────────────────────────────────────
    print("\n▶ Creazione file Excel...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_name in SHEET_ORDER:
        ws = wb.create_sheet(title=sheet_name)
        status  = all_status.get(sheet_name, "error")
        records = all_data.get(sheet_name, [])
        write_sheet(ws, records, portal_status=status)

    wb.save(OUTPUT_PATH)
    print(f"\n✔ File salvato: {OUTPUT_PATH}")

    # ── Riepilogo ─────────────────────────────────────────
    print("\n" + "=" * 65)
    print("RIEPILOGO")
    print("=" * 65)
    for sheet_name in SHEET_ORDER:
        status  = all_status.get(sheet_name, "error")
        n       = len(all_data.get(sheet_name, []))
        flag    = "⚠ NON RAGGIUNGIBILE" if status == "error" else f"{n} atti trovati"
        print(f"  {sheet_name:<25} {flag}")
    print("=" * 65)
    print(f"\nFile: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
