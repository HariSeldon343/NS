#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria - Scraper completo
Portali: ASP CO/CZ/KR/RC/VV, Regione Calabria, TAR Catanzaro
Output: Monitoraggio_Atti_Calabria_DD-MM-YYYY.xlsx
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import json
import time
import datetime
import os
import sys
import traceback
from urllib.parse import urljoin, urlencode, urlparse, parse_qs

# ── Date ─────────────────────────────────────────────────────────────────────
TODAY      = datetime.date.today()           # 2026-05-27
DATE_FROM  = TODAY - datetime.timedelta(days=2)   # 2026-05-25
DATE_FROM_STR  = DATE_FROM.strftime("%Y-%m-%d")   # per ASP e Regione
DATE_FROM_IT   = DATE_FROM.strftime("%d/%m/%Y")   # formato italiano
DATE_TODAY_STR = TODAY.strftime("%Y-%m-%d")
OUTPUT_FNAME   = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
OUTPUT_PATH    = f"/home/user/NS/{OUTPUT_FNAME}"

# ── Keywords ─────────────────────────────────────────────────────────────────
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
# "Life" con word boundary
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)
# Regex per tutte le keyword normali (case-insensitive)
KEYWORDS_RE = [re.compile(re.escape(kw), re.IGNORECASE) for kw in KEYWORDS]

def matches_keywords(text: str) -> bool:
    """Ritorna True se il testo contiene almeno una keyword."""
    if not text:
        return False
    if LIFE_RE.search(text):
        return True
    for pattern in KEYWORDS_RE:
        if pattern.search(text):
            return True
    return False

# ── HTTP Session ──────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
}

def make_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.verify = False  # alcuni portali regionali hanno cert issues
    return s


# ─────────────────────────────────────────────────────────────────────────────
# 1.  SCRAPER ASP (Albo Online SISR Regione Calabria)
# ─────────────────────────────────────────────────────────────────────────────

ASP_PORTALS = [
    ("ASP Cosenza",        "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Catanzaro",      "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Crotone",        "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Reggio Calabria","https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Vibo Valentia",  "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
]

def scrape_asp(name: str, base_url: str) -> list[dict]:
    """Scrape un portale ASP Albo Online e ritorna lista di atti che matchano."""
    results = []
    session = make_session()
    print(f"\n{'='*60}")
    print(f"  Scraping {name}")
    print(f"  URL: {base_url}")

    # Step 1: GET iniziale per ottenere cookie/token di sessione
    try:
        resp = session.get(base_url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"  [ERRORE] GET iniziale: {e}")
        return results

    soup = BeautifulSoup(resp.text, "lxml")

    # Cerca eventuali hidden fields (CSRF, viewstate, ecc.)
    hidden_fields = {}
    for inp in soup.find_all("input", type="hidden"):
        nm = inp.get("name")
        vl = inp.get("value", "")
        if nm:
            hidden_fields[nm] = vl

    print(f"  Hidden fields trovati: {list(hidden_fields.keys())}")

    # Parametri di ricerca standard per questi portali
    # I portali usano tipicamente un form POST con questi campi
    search_params = {
        **hidden_fields,
        "dataPubblicazioneDal": DATE_FROM_STR,
        "dataPubblicazioneAl":  DATE_TODAY_STR,
        "oggetto": "",
        "numero": "",
        "tipo": "",
    }

    # Cerca il form e l'action URL
    form = soup.find("form")
    action_url = base_url
    if form:
        action = form.get("action", "")
        if action:
            action_url = urljoin(base_url, action)
        form_method = form.get("method", "get").lower()
        print(f"  Form action: {action_url}, method: {form_method}")
        # Raccoglie tutti gli input del form
        for inp in form.find_all(["input", "select"]):
            nm  = inp.get("name")
            vl  = inp.get("value", "")
            typ = inp.get("type", "text")
            if nm and nm not in hidden_fields:
                if typ not in ("submit", "button", "image", "reset"):
                    search_params[nm] = vl

    print(f"  Ricerca dal: {DATE_FROM_STR} al: {DATE_TODAY_STR}")

    # Paginazione
    page = 1
    total_fetched = 0

    while True:
        print(f"  Pagina {page}...", end=" ")

        # Aggiunge parametro pagina
        params = {**search_params}
        if page > 1:
            # Prova diversi nomi comuni per il parametro pagina
            params["page"]          = page
            params["pagina"]        = page
            params["currentPage"]   = page
            params["paginaCorrente"]= page

        try:
            if form and form.get("method", "get").lower() == "post":
                resp = session.post(action_url, data=params, timeout=30)
            else:
                resp = session.get(action_url, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"[ERRORE] {e}")
            break

        soup = BeautifulSoup(resp.text, "lxml")

        # Cerca la tabella dei risultati
        table = (
            soup.find("table", id=re.compile(r"result|albo|atti", re.I))
            or soup.find("table", class_=re.compile(r"result|albo|atti|table", re.I))
            or soup.find("table")
        )

        if not table:
            print(f"[nessuna tabella]")
            if page == 1:
                # Debug: salva HTML
                debug_file = f"/tmp/asp_debug_{name.replace(' ','_')}_p{page}.html"
                with open(debug_file, "w", encoding="utf-8") as f:
                    f.write(resp.text)
                print(f"  [DEBUG] HTML salvato in {debug_file}")
            break

        # Parse righe tabella
        rows = table.find_all("tr")
        headers = []
        data_rows = []
        for i, row in enumerate(rows):
            cells = row.find_all(["th", "td"])
            if i == 0 or row.find("th"):
                headers = [c.get_text(strip=True) for c in cells]
            else:
                data_rows.append(cells)

        print(f"{len(data_rows)} righe", end=" ")

        if not data_rows:
            print()
            break

        page_results = 0
        for row_cells in data_rows:
            # Estrai testo e link dalle celle
            cell_texts = [c.get_text(strip=True) for c in row_cells]
            link_tag   = None
            link_url   = ""

            # Cerca link nella riga
            for cell in row_cells:
                a = cell.find("a", href=True)
                if a:
                    link_tag = a
                    link_url = urljoin(base_url, a["href"])
                    break

            # Ricostruisce un dizionario con le colonne
            row_dict = {}
            for j, txt in enumerate(cell_texts):
                key = headers[j] if j < len(headers) else f"col{j}"
                row_dict[key] = txt

            # Cerca l'oggetto (vari nomi possibili)
            oggetto = (
                row_dict.get("Oggetto")
                or row_dict.get("oggetto")
                or row_dict.get("Descrizione")
                or row_dict.get("Titolo")
                or " ".join(cell_texts[2:4]) if len(cell_texts) >= 4 else " ".join(cell_texts)
            )

            # Cerca la data (vari nomi possibili)
            data_atto = (
                row_dict.get("Data pubblicazione")
                or row_dict.get("Data Pubblicazione")
                or row_dict.get("Data")
                or row_dict.get("data")
                or row_dict.get("DataPubblicazione")
                or (cell_texts[0] if cell_texts else "")
            )

            if matches_keywords(oggetto):
                results.append({
                    "data":     data_atto,
                    "oggetto":  oggetto,
                    "link":     link_url,
                    "raw":      row_dict,
                })
                page_results += 1

        total_fetched += len(data_rows)
        print(f"→ {page_results} match")

        # Controlla se c'è una pagina successiva
        next_page = _find_next_page(soup, page)
        if next_page is None:
            break
        page = next_page
        time.sleep(0.5)

    print(f"  Totale righe analizzate: {total_fetched}, match: {len(results)}")
    return results


def _find_next_page(soup: BeautifulSoup, current_page: int) -> int | None:
    """
    Cerca un link/pulsante 'pagina successiva'.
    Ritorna il numero della prossima pagina o None se non c'è.
    """
    # Cerca link con testo ">" o "Successiva" o "Next"
    for a in soup.find_all("a"):
        txt = a.get_text(strip=True)
        if txt in (">", "»", "Successiva", "Next", "Avanti"):
            href = a.get("href", "")
            if href and href != "#":
                # Prova a estrarre numero pagina dall'href
                m = re.search(r"[pP]age[=_]?(\d+)|[pP]agina[=_]?(\d+)|currentPage[=]?(\d+)", href)
                if m:
                    pg = next(int(x) for x in m.groups() if x)
                    return pg
                return current_page + 1

    # Cerca pattern di paginazione con numeri
    pagination = soup.find(class_=re.compile(r"paginat|pager|pagination", re.I))
    if pagination:
        for a in pagination.find_all("a"):
            txt = a.get_text(strip=True)
            if txt.isdigit() and int(txt) == current_page + 1:
                return current_page + 1

    return None


# ─────────────────────────────────────────────────────────────────────────────
# 2.  SCRAPER REGIONE CALABRIA
# ─────────────────────────────────────────────────────────────────────────────

REGIONE_BASE = "https://www.regione.calabria.it/provvedimenti-della-regione/"

def scrape_regione() -> list[dict]:
    results = []
    session = make_session()
    print(f"\n{'='*60}")
    print(f"  Scraping Regione Calabria")
    print(f"  URL: {REGIONE_BASE}")

    page = 1
    total_fetched = 0

    while True:
        print(f"  Pagina {page}...", end=" ")

        # Il sito Regione Calabria usa tipicamente WordPress con query params
        params = {
            "filter_date_from": DATE_FROM_STR,
            "filter_date_to":   DATE_TODAY_STR,
            "paged":            page,
        }
        if page == 1:
            del params["paged"]

        try:
            resp = session.get(REGIONE_BASE, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"[ERRORE] {e}")
            break

        soup = BeautifulSoup(resp.text, "lxml")

        # Debug pagina 1
        if page == 1:
            debug_file = "/tmp/regione_p1.html"
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(resp.text)
            print(f"  [DEBUG] HTML salvato in {debug_file}")

        # Cerca tabella o lista risultati
        rows_data = _parse_regione_page(soup, REGIONE_BASE)
        print(f"{len(rows_data)} righe", end=" ")

        if not rows_data:
            break

        page_match = 0
        for row in rows_data:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)
                page_match += 1

        total_fetched += len(rows_data)
        print(f"→ {page_match} match")

        # Cerca next page
        if not _has_next_page_regione(soup):
            break
        page += 1
        time.sleep(0.5)

    print(f"  Totale righe analizzate: {total_fetched}, match: {len(results)}")
    return results


def _parse_regione_page(soup: BeautifulSoup, base_url: str) -> list[dict]:
    items = []

    # Prova tabella
    table = soup.find("table")
    if table:
        rows = table.find_all("tr")
        headers = []
        for i, row in enumerate(rows):
            cells = row.find_all(["th", "td"])
            if i == 0 or row.find("th"):
                headers = [c.get_text(strip=True) for c in cells]
                continue
            cell_texts = [c.get_text(strip=True) for c in cells]
            link_url = ""
            for cell in cells:
                a = cell.find("a", href=True)
                if a:
                    link_url = urljoin(base_url, a["href"])
                    break
            row_dict = {headers[j] if j < len(headers) else f"col{j}": t
                        for j, t in enumerate(cell_texts)}
            oggetto  = (row_dict.get("Oggetto") or row_dict.get("Descrizione")
                        or row_dict.get("Titolo") or cell_texts[1] if len(cell_texts) > 1 else "")
            data_row = (row_dict.get("Data") or row_dict.get("data")
                        or cell_texts[0] if cell_texts else "")
            items.append({"data": data_row, "oggetto": oggetto, "link": link_url, "raw": row_dict})
        return items

    # Prova lista articoli (WordPress style)
    articles = soup.find_all(["article", "li"], class_=re.compile(r"post|item|row|result", re.I))
    for art in articles:
        a = art.find("a", href=True)
        title = art.get_text(strip=True)[:300]
        link  = urljoin(base_url, a["href"]) if a else ""
        # Cerca data
        date_tag = art.find(class_=re.compile(r"date|data", re.I))
        data_row = date_tag.get_text(strip=True) if date_tag else ""
        items.append({"data": data_row, "oggetto": title, "link": link})

    return items


def _has_next_page_regione(soup: BeautifulSoup) -> bool:
    for a in soup.find_all("a"):
        txt = a.get_text(strip=True)
        if txt in (">", "»", "Successiva", "Next", "Avanti", "›"):
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# 3.  SCRAPER TAR CATANZARO
# ─────────────────────────────────────────────────────────────────────────────

TAR_BASE = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
# Per TAR: ultimi 3 giorni
DATE_FROM_TAR = TODAY - datetime.timedelta(days=3)
DATE_FROM_TAR_STR = DATE_FROM_TAR.strftime("%Y-%m-%d")

def scrape_tar() -> list[dict]:
    results = []
    session = make_session()
    print(f"\n{'='*60}")
    print(f"  Scraping TAR Catanzaro")
    print(f"  URL: {TAR_BASE}")
    print(f"  Ricerca dal: {DATE_FROM_TAR_STR}")

    page = 0  # Spesso 0-indexed
    total_fetched = 0

    while True:
        print(f"  Pagina {page+1}...", end=" ")

        params = {
            "publishDateFrom": DATE_FROM_TAR_STR,
            "publishDateTo":   DATE_TODAY_STR,
            "page":            page,
            "size":            50,
        }

        try:
            resp = session.get(TAR_BASE, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"[ERRORE] {e}")
            break

        soup = BeautifulSoup(resp.text, "lxml")

        if page == 0:
            debug_file = "/tmp/tar_p1.html"
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(resp.text)
            print(f"  [DEBUG] HTML salvato in {debug_file}")

        rows_data = _parse_tar_page(soup, TAR_BASE)
        print(f"{len(rows_data)} righe", end=" ")

        if not rows_data:
            break

        page_match = 0
        for row in rows_data:
            # Filtra per keyword su "Parte" e "Oggetto"
            parte   = row.get("parte", "")
            oggetto = row.get("oggetto", "")
            combined = f"{parte} {oggetto}"
            if matches_keywords(combined):
                results.append(row)
                page_match += 1

        total_fetched += len(rows_data)
        print(f"→ {page_match} match")

        if not _has_next_page_tar(soup):
            break
        page += 1
        time.sleep(0.5)

    print(f"  Totale righe analizzate: {total_fetched}, match: {len(results)}")
    return results


def _parse_tar_page(soup: BeautifulSoup, base_url: str) -> list[dict]:
    items = []

    table = soup.find("table")
    if not table:
        # Prova ricerca più ampia
        table = soup.find(class_=re.compile(r"result|provvediment|table", re.I))

    if table:
        rows = table.find_all("tr")
        headers = []
        for i, row in enumerate(rows):
            cells = row.find_all(["th", "td"])
            if i == 0 or row.find("th"):
                headers = [c.get_text(strip=True) for c in cells]
                continue
            cell_texts = [c.get_text(strip=True) for c in cells]
            link_url = ""
            for cell in cells:
                a = cell.find("a", href=True)
                if a:
                    link_url = urljoin(base_url, a["href"])
                    break

            row_dict = {headers[j] if j < len(headers) else f"col{j}": t
                        for j, t in enumerate(cell_texts)}

            parte   = (row_dict.get("Parte") or row_dict.get("parte")
                       or row_dict.get("Ricorrente") or "")
            oggetto = (row_dict.get("Oggetto") or row_dict.get("oggetto")
                       or row_dict.get("Materia") or "")
            data_row= (row_dict.get("Data")   or row_dict.get("data")
                       or row_dict.get("DataDeposito") or "")

            items.append({
                "data":    data_row,
                "parte":   parte,
                "oggetto": oggetto,
                "link":    link_url,
                "raw":     row_dict,
            })

    return items


def _has_next_page_tar(soup: BeautifulSoup) -> bool:
    for a in soup.find_all("a"):
        txt = a.get_text(strip=True)
        if txt in (">", "»", "Successiva", "Next", "Avanti", "›"):
            href = a.get("href", "")
            if href and href != "#":
                return True
    # Cerca paginazione
    pag = soup.find(class_=re.compile(r"paginat|pager|pagination", re.I))
    if pag:
        for a in pag.find_all("a"):
            txt = a.get_text(strip=True)
            if txt in (">", "»", "›"):
                return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# 4.  GENERAZIONE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

HYPERLINK_FONT = Font(color="0563C1", underline="single")
HEADER_FILL    = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT    = Font(bold=True, color="FFFFFF", size=11)
ALT_FILL       = PatternFill("solid", fgColor="DCE6F1")

def write_sheet(ws, rows: list[dict], is_tar: bool = False):
    """Scrive un foglio Excel con i risultati."""

    # Header
    if is_tar:
        headers = ["Data", "Parte", "Oggetto", "Descrizione breve (150 car.)", "Link"]
    else:
        headers = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font    = HEADER_FONT
        cell.fill    = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 30

    if not rows:
        ws.cell(row=2, column=1, value="Nessun risultato")
        ws.column_dimensions["A"].width = 30
        return

    for r_idx, row in enumerate(rows, 2):
        fill = ALT_FILL if r_idx % 2 == 0 else PatternFill()

        data_val   = row.get("data", "")
        oggetto    = row.get("oggetto", "") or ""
        parte      = row.get("parte", "")  or ""
        descr      = (oggetto[:150] + "…") if len(oggetto) > 150 else oggetto
        link_url   = row.get("link", "") or ""

        if is_tar:
            vals = [data_val, parte, oggetto, descr, link_url]
        else:
            vals = [data_val, oggetto, descr, link_url]

        link_col = 5 if is_tar else 4

        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=r_idx, column=col, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=(col == (3 if is_tar else 2)))
            if fill.fill_type:
                cell.fill = fill
            # Link cliccabile
            if col == link_col and link_url:
                cell.value     = link_url
                cell.hyperlink = link_url
                cell.font      = HYPERLINK_FONT

    # Larghezze colonne
    if is_tar:
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 60
        ws.column_dimensions["D"].width = 40
        ws.column_dimensions["E"].width = 60
    else:
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 65
        ws.column_dimensions["C"].width = 45
        ws.column_dimensions["D"].width = 60

    # Freeze header
    ws.freeze_panes = "A2"


# ─────────────────────────────────────────────────────────────────────────────
# 5.  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print(f"\n{'#'*60}")
    print(f"  MONITORAGGIO ATTI CALABRIA")
    print(f"  Data esecuzione: {TODAY.strftime('%d/%m/%Y')}")
    print(f"  Periodo: {DATE_FROM_STR} → {DATE_TODAY_STR}")
    print(f"  Output: {OUTPUT_PATH}")
    print(f"{'#'*60}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuovi foglio default

    # ── ASP Portali ──────────────────────────────────────────────────────────
    asp_sheet_results = {}
    for asp_name, asp_url in ASP_PORTALS:
        try:
            data = scrape_asp(asp_name, asp_url)
        except Exception as e:
            print(f"  [ERRORE FATALE] {asp_name}: {e}")
            traceback.print_exc()
            data = []
        asp_sheet_results[asp_name] = data

    for asp_name in [n for n, _ in ASP_PORTALS]:
        ws = wb.create_sheet(title=asp_name[:31])
        write_sheet(ws, asp_sheet_results.get(asp_name, []))
        print(f"\n  Foglio '{asp_name}': {len(asp_sheet_results.get(asp_name, []))} righe")

    # ── Regione Calabria ─────────────────────────────────────────────────────
    try:
        regione_data = scrape_regione()
    except Exception as e:
        print(f"  [ERRORE FATALE] Regione Calabria: {e}")
        traceback.print_exc()
        regione_data = []

    ws_reg = wb.create_sheet(title="Regione Calabria")
    write_sheet(ws_reg, regione_data)
    print(f"\n  Foglio 'Regione Calabria': {len(regione_data)} righe")

    # ── TAR Catanzaro ─────────────────────────────────────────────────────────
    try:
        tar_data = scrape_tar()
    except Exception as e:
        print(f"  [ERRORE FATALE] TAR Catanzaro: {e}")
        traceback.print_exc()
        tar_data = []

    ws_tar = wb.create_sheet(title="TAR Catanzaro")
    write_sheet(ws_tar, tar_data, is_tar=True)
    print(f"\n  Foglio 'TAR Catanzaro': {len(tar_data)} righe")

    # ── Salva ────────────────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"\n{'#'*60}")
    print(f"  FILE SALVATO: {OUTPUT_PATH}")
    print(f"{'#'*60}\n")


if __name__ == "__main__":
    main()
