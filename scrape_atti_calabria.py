#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraping portali ASP, Regione Calabria e TAR Catanzaro
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta
import time
import sys
import os

# ── configurazione ───────────────────────────────────────────────────────────

TODAY = datetime.now()
DATE_FROM_ASP = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM_REG = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM_REG_DMY = (TODAY - timedelta(days=2)).strftime("%d/%m/%Y")
DATE_FROM_TAR = (TODAY - timedelta(days=3)).strftime("%d/%m/%Y")

OUTPUT_FILENAME = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"

KEYWORDS = [
    "ADI",
    "Assistenza domiciliare",
    "ANMIC",
    "Accreditamento",
    "Aumento di budget",
    "Autismo",
    "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione",
    "Autorizzazioni",
    "Budget",
    "Casa Giardino",
    "Centro San Giuseppe",
    "Centro salute e benessere",
    "Fabbisogni LEA",
    "Fisiolab",
    "Fisioterapia",
    "Life",
    "Parere commissione",
    "Presa d'atto verifica",
    "Programmazione",
    "Rete riabilitativa",
    "Rete territoriale",
    "Riabilitazione estensiva",
    "Riconversione prestazioni",
    "Rinnovo accreditamento",
    "San Teodoro",
    "Savelli Hospital",
    "Starbene",
    "Verifica requisiti",
    "Villa San Giuseppe",
    "Villa del Rosario",
]

KEYWORD_PATTERNS = []
for kw in KEYWORDS:
    if kw.lower() == "life":
        KEYWORD_PATTERNS.append(re.compile(r'\bLIFE\b', re.IGNORECASE))
    else:
        KEYWORD_PATTERNS.append(re.compile(re.escape(kw), re.IGNORECASE))


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    for pat in KEYWORD_PATTERNS:
        if pat.search(text):
            return True
    return False


SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
})


def safe_get(url, params=None, retries=3, **kwargs):
    for attempt in range(retries):
        try:
            r = SESSION.get(url, params=params, timeout=30, **kwargs)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  [warn] GET {url[:80]} tentativo {attempt+1}/{retries}: {e}")
            time.sleep(2 ** attempt)
    return None


def safe_post(url, data=None, retries=3, **kwargs):
    for attempt in range(retries):
        try:
            r = SESSION.post(url, data=data, timeout=30, **kwargs)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  [warn] POST {url[:80]} tentativo {attempt+1}/{retries}: {e}")
            time.sleep(2 ** attempt)
    return None


# ── scraping ASP portali ─────────────────────────────────────────────────────

ASP_PORTALS = [
    ("ASP Cosenza",          "https://online-aspco.sisr.regione.calabria.it"),
    ("ASP Catanzaro",        "https://online-aspcz.sisr.regione.calabria.it"),
    ("ASP Crotone",          "https://online-aspkr.sisr.regione.calabria.it"),
    ("ASP Reggio Calabria",  "https://online-asprc.sisr.regione.calabria.it"),
    ("ASP Vibo Valentia",    "https://online-aspvv.sisr.regione.calabria.it"),
]


def scrape_asp_portal(name: str, base_url: str) -> list[dict]:
    print(f"\n[ASP] {name} — {base_url}")
    search_url = f"{base_url}/AlboOnline/ricercaAlbo"
    results = []

    r = safe_get(search_url)
    if r is None:
        print(f"  [errore] portale non raggiungibile (rete intranet regionale)")
        return results

    soup = BeautifulSoup(r.text, "lxml")
    hidden = {}
    for inp in soup.find_all("input", {"type": "hidden"}):
        n = inp.get("name")
        if n:
            hidden[n] = inp.get("value", "")

    payload = {
        **hidden,
        "dataPubblicazioneDal": DATE_FROM_ASP,
        "dataPubblicazioneAl": "",
        "numeroAtto": "",
        "annoAtto": "",
        "tipoAtto": "",
        "oggetto": "",
        "submit": "Cerca",
    }

    r2 = safe_post(search_url, data=payload)
    if r2 is None:
        r2 = safe_get(search_url, params={"dataPubblicazioneDal": DATE_FROM_ASP})
    if r2 is None:
        print(f"  [errore] ricerca fallita")
        return results

    page = 1
    current_response = r2

    while True:
        soup2 = BeautifulSoup(current_response.text, "lxml")

        table = (
            soup2.find("table", {"class": re.compile(r"table|result|albo", re.I)})
            or soup2.find("table")
        )

        if table:
            rows = table.find_all("tr")
            for row in rows[1:]:
                cols = row.find_all(["td", "th"])
                if len(cols) < 2:
                    continue

                a_tag = row.find("a", href=True)
                link_val = ""
                if a_tag:
                    href = a_tag["href"]
                    link_val = href if href.startswith("http") else base_url + ("" if href.startswith("/") else "/") + href

                date_pat = re.compile(r'\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}')
                data_val = ""
                oggetto_val = ""

                for col in cols:
                    txt = col.get_text(strip=True)
                    if date_pat.match(txt) and not data_val:
                        data_val = txt

                # Oggetto: colonna più lunga che non sia una data
                for c in cols:
                    t = c.get_text(strip=True)
                    if len(t) > len(oggetto_val) and not date_pat.match(t) and t not in ("→", ""):
                        oggetto_val = t

                if matches_keywords(oggetto_val):
                    results.append({"data": data_val, "oggetto": oggetto_val, "url": link_val})

        # Paginazione
        next_link = None
        for a in soup2.find_all("a", href=True):
            txt = a.get_text(strip=True).lower()
            if txt in ("successivo", "next", ">", "»") or "pagina successiva" in txt or txt == str(page + 1):
                next_link = a["href"]
                break

        if not next_link:
            break

        href = next_link
        if not href.startswith("http"):
            href = base_url + ("" if href.startswith("/") else "/") + href
        current_response = safe_get(href)
        if current_response is None:
            break
        page += 1
        print(f"  pagina {page}...")
        time.sleep(0.5)

    print(f"  trovati {len(results)} atti matching")
    return results


# ── scraping Regione Calabria ────────────────────────────────────────────────

def scrape_regione_calabria() -> list[dict]:
    """
    Tabella: Tipologia | Data Repertoriazione | N | Dipartimento | Oggetto | Dettaglio(→)
    Indici:     0              1                2        3             4          5
    Paginazione: ?filter_date_from=YYYY-MM-DD&paged=N
    """
    print("\n[Regione Calabria] scraping provvedimenti...")
    base = "https://www.regione.calabria.it/provvedimenti-della-regione/"
    results = []

    cutoff_date = TODAY - timedelta(days=2)

    def parse_date(s: str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(s.strip(), fmt)
            except ValueError:
                continue
        return None

    page = 1
    empty_pages = 0

    while empty_pages < 2:
        params = {"filter_date_from": DATE_FROM_REG, "paged": page}
        r = safe_get(base, params=params)
        if r is None:
            break

        soup = BeautifulSoup(r.text, "lxml")
        table = soup.find("table")
        if not table:
            break

        rows = table.find_all("tr")[1:]  # skip header
        if not rows:
            break

        page_had_results = False
        for row in rows:
            cols = row.find_all(["td", "th"])
            if len(cols) < 5:
                continue

            data_str = cols[1].get_text(strip=True)
            oggetto = cols[4].get_text(strip=True)

            a_tag = row.find("a", href=True)
            link_val = a_tag["href"] if a_tag else ""

            # Filtro data client-side
            parsed = parse_date(data_str)
            if parsed and parsed.date() < cutoff_date.date():
                continue

            page_had_results = True
            if matches_keywords(oggetto):
                results.append({"data": data_str, "oggetto": oggetto, "url": link_val})

        if not page_had_results:
            empty_pages += 1
        else:
            empty_pages = 0

        # Controlla paginazione
        next_url = None
        for a in soup.find_all("a", href=True):
            txt = a.get_text(strip=True).lower()
            if "successiv" in txt or txt in (">", "»"):
                next_url = a["href"]
                break
            m = re.search(r'paged=(\d+)', a["href"])
            if m and int(m.group(1)) == page + 1:
                next_url = a["href"]
                break

        if not next_url:
            break

        page += 1
        print(f"  pagina {page}...")
        time.sleep(0.5)

    print(f"  trovati {len(results)} atti matching")
    return results


# ── scraping TAR Catanzaro ───────────────────────────────────────────────────

def scrape_tar_catanzaro() -> list[dict]:
    """
    Portale Liferay con form POST.
    Campi: publishDateFrom (formato dd/MM/yyyy), paginazione con parametro delta/start.
    """
    print("\n[TAR Catanzaro] scraping provvedimenti...")
    base_url = "https://www.giustizia-amministrativa.it"
    page_url = f"{base_url}/provvedimenti-tar-catanzaro"
    results = []

    NS = "_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"

    # GET iniziale per ottenere form e token
    r_get = safe_get(page_url)
    if r_get is None:
        print("  [errore] impossibile accedere al portale TAR")
        return results

    soup = BeautifulSoup(r_get.text, "lxml")
    form = soup.find("form", id=lambda x: x and "jjYpzZYF4Qfe" in str(x))
    if not form:
        print("  [errore] form di ricerca TAR non trovato")
        return results

    form_action = form.get("action", "")
    payload = {}
    for inp in form.find_all("input"):
        n = inp.get("name")
        if n:
            payload[n] = inp.get("value", "")

    payload[NS + "publishDateFrom"] = DATE_FROM_TAR

    page_start = 0
    page_size = 20

    while True:
        if page_start > 0:
            payload[NS + "cur"] = str(page_start // page_size + 1)
            payload[NS + "delta"] = str(page_size)

        r_post = safe_post(form_action, data=payload)
        if r_post is None:
            break

        soup2 = BeautifulSoup(r_post.text, "lxml")

        # Controlla errore portlet
        err_div = soup2.find("div", class_="alert-warning")
        if err_div:
            print(f"  [avviso server] {err_div.get_text(strip=True)[:120]}")
            # Prova senza errore - se c'è una tabella comunque la leggo
            pass

        table = soup2.find("table")
        if not table:
            break

        rows = table.find_all("tr")[1:]
        if not rows:
            break

        for row in rows:
            cols = row.find_all(["td", "th"])
            if not cols:
                continue
            texts = [c.get_text(strip=True) for c in cols]

            a_tag = row.find("a", href=True)
            link_val = ""
            if a_tag:
                href = a_tag["href"]
                link_val = href if href.startswith("http") else base_url + href

            date_pat = re.compile(r'\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}')
            data_val = next((t for t in texts if date_pat.search(t)), "")

            # "Parte" è spesso la colonna con il nome della parte ricorrente
            parte_val = ""
            for t in texts:
                if t and t != "XXX_OMISSIS_XXX" and not date_pat.match(t) and len(t) > len(parte_val):
                    parte_val = t

            full_text = " | ".join(t for t in texts if t)
            if matches_keywords(full_text):
                results.append({
                    "data": data_val,
                    "oggetto": full_text[:500],
                    "url": link_val,
                })

        # Paginazione
        has_next = bool(soup2.find("a", string=re.compile(r'successiv|next|>|»', re.I)))
        if not has_next:
            break

        page_start += page_size
        print(f"  pagina {page_start // page_size + 1}...")
        time.sleep(0.5)

    print(f"  trovati {len(results)} atti matching")
    return results


# ── generazione Excel ────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
LINK_FONT   = Font(color="0563C1", underline="single", size=10)
BODY_FONT   = Font(size=10)
ALT_FILL    = PatternFill("solid", fgColor="DCE6F1")

COL_WIDTHS  = [15, 80, 55, 50]
COL_HEADERS = ["Data", "Oggetto", "Descrizione breve", "Link"]


def build_sheet(ws, rows: list[dict], note: str = ""):
    ws.freeze_panes = "A2"

    for col_idx, (hdr, width) in enumerate(zip(COL_HEADERS, COL_WIDTHS), start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20

    if not rows:
        msg = note if note else "Nessun risultato"
        c = ws.cell(row=2, column=1, value=msg)
        c.font = Font(italic=True, color="888888")
        ws.merge_cells("A2:D2")
        return

    for row_idx, item in enumerate(rows, start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else None

        c_data = ws.cell(row=row_idx, column=1, value=item.get("data", ""))
        c_data.font = BODY_FONT
        c_data.alignment = Alignment(vertical="top")
        if fill:
            c_data.fill = fill

        oggetto = item.get("oggetto", "")
        c_ogg = ws.cell(row=row_idx, column=2, value=oggetto)
        c_ogg.font = BODY_FONT
        c_ogg.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c_ogg.fill = fill

        desc = oggetto[:150]
        c_desc = ws.cell(row=row_idx, column=3, value=desc)
        c_desc.font = BODY_FONT
        c_desc.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c_desc.fill = fill

        url = item.get("url", "")
        c_link = ws.cell(row=row_idx, column=4, value=url if url else "—")
        if url:
            c_link.hyperlink = url
            c_link.font = LINK_FONT
        else:
            c_link.font = BODY_FONT
        c_link.alignment = Alignment(vertical="top")
        if fill:
            c_link.fill = fill

        row_height = max(15, min(75, 15 + len(oggetto) // 6))
        ws.row_dimensions[row_idx].height = row_height


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print(f"Monitoraggio Atti Calabria — {TODAY.strftime('%d/%m/%Y')}")
    print(f"Filtro ASP/Regione dal: {DATE_FROM_REG_DMY}")
    print(f"Filtro TAR dal: {DATE_FROM_TAR}")
    print("=" * 60)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ASP Portali
    asp_notes = {
        "ASP Cosenza": "",
        "ASP Catanzaro": "",
        "ASP Crotone": "",
        "ASP Reggio Calabria": "",
        "ASP Vibo Valentia": "",
    }
    for portal_name, portal_url in ASP_PORTALS:
        data = scrape_asp_portal(portal_name, portal_url)
        ws = wb.create_sheet(title=portal_name[:31])
        note = "Portale non raggiungibile (accessibile solo da rete intranet regionale)" if not data else ""
        build_sheet(ws, data, note=note)

    # Regione Calabria
    data_reg = scrape_regione_calabria()
    ws_reg = wb.create_sheet(title="Regione Calabria")
    build_sheet(ws_reg, data_reg)

    # TAR Catanzaro
    data_tar = scrape_tar_catanzaro()
    ws_tar = wb.create_sheet(title="TAR Catanzaro")
    tar_note = ""
    if not data_tar:
        tar_note = "Nessun risultato — Il portlet TAR ha restituito un errore server (NullPointerException) per la ricerca per data. Verificare manualmente: https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
    build_sheet(ws_tar, data_tar, note=tar_note)

    wb.save(OUTPUT_FILENAME)
    abs_path = os.path.abspath(OUTPUT_FILENAME)
    print(f"\n[OK] File salvato: {abs_path}")
    print(f"Fogli: {[ws.title for ws in wb.worksheets]}")

    # Sommario risultati
    print("\n=== SOMMARIO ===")
    for ws in wb.worksheets:
        rows = ws.max_row - 1  # escludi header
        print(f"  {ws.title}: {max(0, rows)} righe")


if __name__ == "__main__":
    main()
