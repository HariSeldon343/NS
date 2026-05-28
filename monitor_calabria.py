#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Accede ai portali ASP, Regione Calabria e TAR Catanzaro,
filtra per keyword e produce un file Excel con hyperlink.
"""

import re
import time
import logging
from datetime import date, timedelta
from urllib.parse import urljoin, urlencode, quote

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── date ──────────────────────────────────────────────────────────────────────
TODAY = date(2026, 5, 28)
DATE_FROM_2D = TODAY - timedelta(days=2)  # 2026-05-26
DATE_FROM_3D = TODAY - timedelta(days=3)  # 2026-05-25

# ── keywords ──────────────────────────────────────────────────────────────────
KEYWORDS_NORMAL = [
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
# "Life" usa word boundary per evitare falsi positivi
KEYWORD_LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

# Build combined regex
_parts = [re.escape(kw) for kw in KEYWORDS_NORMAL]
KEYWORDS_RE = re.compile('|'.join(_parts), re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    return bool(KEYWORDS_RE.search(text)) or bool(KEYWORD_LIFE_RE.search(text))


# ── HTTP session ───────────────────────────────────────────────────────────────
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
})


def safe_get(url, params=None, timeout=30, retries=3, **kwargs):
    for attempt in range(retries):
        try:
            r = SESSION.get(url, params=params, timeout=timeout, **kwargs)
            r.raise_for_status()
            return r
        except Exception as exc:
            log.warning("GET %s attempt %d/%d failed: %s", url, attempt + 1, retries, exc)
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    return None


def safe_post(url, data=None, timeout=30, retries=3, **kwargs):
    for attempt in range(retries):
        try:
            r = SESSION.post(url, data=data, timeout=timeout, **kwargs)
            r.raise_for_status()
            return r
        except Exception as exc:
            log.warning("POST %s attempt %d/%d failed: %s", url, attempt + 1, retries, exc)
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    return None


# ── ASP portals ───────────────────────────────────────────────────────────────
ASP_PORTALS = [
    ("ASP Cosenza",          "https://online-aspco.sisr.regione.calabria.it"),
    ("ASP Catanzaro",        "https://online-aspcz.sisr.regione.calabria.it"),
    ("ASP Crotone",          "https://online-aspkr.sisr.regione.calabria.it"),
    ("ASP Reggio Calabria",  "https://online-asprc.sisr.regione.calabria.it"),
    ("ASP Vibo Valentia",    "https://online-aspvv.sisr.regione.calabria.it"),
]

SEARCH_PATH = "/AlboOnline/ricercaAlbo"


def fetch_asp_portal(name: str, base_url: str) -> list[dict]:
    """Fetch all matching acts from an ASP portal."""
    results = []
    search_url = base_url + SEARCH_PATH
    date_from_str = DATE_FROM_2D.strftime("%Y-%m-%d")

    log.info("[%s] Starting fetch from %s", name, search_url)

    # First: GET the search page to get any CSRF token / form fields
    r0 = safe_get(search_url)
    if r0 is None:
        log.error("[%s] Cannot reach search page", name)
        return results

    soup0 = BeautifulSoup(r0.text, "lxml")

    # Collect hidden form fields
    form = soup0.find("form")
    hidden_fields = {}
    if form:
        for inp in form.find_all("input", type="hidden"):
            if inp.get("name"):
                hidden_fields[inp["name"]] = inp.get("value", "")

    page = 1
    while True:
        log.info("[%s] Fetching page %d", name, page)
        payload = {
            **hidden_fields,
            "dataPubblicazioneDal": date_from_str,
            "page": str(page),
        }

        r = safe_post(search_url, data=payload)
        if r is None:
            log.error("[%s] Failed to fetch page %d", name, page)
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows = _parse_asp_rows(soup, base_url)

        if not rows:
            # Try GET with params as fallback
            if page == 1:
                params = {"dataPubblicazioneDal": date_from_str}
                r2 = safe_get(search_url, params=params)
                if r2:
                    soup2 = BeautifulSoup(r2.text, "lxml")
                    rows = _parse_asp_rows(soup2, base_url)
            if not rows:
                log.info("[%s] No rows on page %d – stopping", name, page)
                break

        for row in rows:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)

        # Check for next page
        if not _has_next_page(soup, page):
            break
        page += 1

    log.info("[%s] Found %d matching acts", name, len(results))
    return results


def _parse_asp_rows(soup: BeautifulSoup, base_url: str) -> list[dict]:
    rows = []
    table = soup.find("table", {"id": re.compile(r"risultati|atti|albo", re.I)})
    if not table:
        table = soup.find("table")
    if not table:
        return rows

    headers = []
    header_row = table.find("thead")
    if header_row:
        headers = [th.get_text(strip=True).lower() for th in header_row.find_all(["th", "td"])]

    tbody = table.find("tbody") or table
    for tr in tbody.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if not cells or len(cells) < 2:
            continue

        # Skip header rows
        if all(c.name == "th" for c in cells):
            continue

        row_data = {}
        cell_texts = [c.get_text(strip=True) for c in cells]

        # Try to map by headers
        if headers and len(headers) == len(cells):
            for h, c in zip(headers, cells):
                row_data[h] = c.get_text(strip=True)
        else:
            # Fallback: assume standard layout
            # Common: numero, tipo, data, oggetto, link
            for i, (h, c) in enumerate(zip(
                ["numero", "tipo", "data", "oggetto", "extra"], cells
            )):
                row_data[h] = c.get_text(strip=True)

        # Find link
        link = None
        for a in tr.find_all("a", href=True):
            href = a["href"]
            if href and href != "#":
                link = urljoin(base_url, href)
                break

        # Determine oggetto and data
        oggetto = (
            row_data.get("oggetto")
            or row_data.get("descrizione")
            or row_data.get("titolo")
            or ""
        )
        data_pubbl = (
            row_data.get("data")
            or row_data.get("data pubblicazione")
            or row_data.get("data_pubblicazione")
            or row_data.get("data pubbl.")
            or ""
        )

        if oggetto or link:
            rows.append({
                "data": data_pubbl,
                "oggetto": oggetto,
                "link": link or "",
                "_raw": cell_texts,
            })

    return rows


def _has_next_page(soup: BeautifulSoup, current_page: int) -> bool:
    # Look for pagination links
    pag = soup.find(class_=re.compile(r"paginat|pagination|pager", re.I))
    if not pag:
        return False
    # Find a link to the next page number
    for a in pag.find_all("a", href=True):
        txt = a.get_text(strip=True)
        if txt == str(current_page + 1):
            return True
        if re.search(r"next|successiv[oa]|>>|›", txt, re.I):
            return True
    return False


# ── Regione Calabria ──────────────────────────────────────────────────────────
RC_BASE = "https://www.regione.calabria.it"
RC_PATH = "/provvedimenti-della-regione/"


def fetch_regione_calabria() -> list[dict]:
    results = []
    date_from_str = DATE_FROM_2D.strftime("%Y-%m-%d")
    base_url = RC_BASE + RC_PATH

    log.info("[Regione Calabria] Starting fetch from %s", base_url)

    page = 1
    while True:
        log.info("[Regione Calabria] Fetching page %d", page)
        params = {
            "filter_date_from": date_from_str,
            "page": str(page),
        }
        r = safe_get(base_url, params=params)
        if r is None:
            log.error("[Regione Calabria] Failed to fetch page %d", page)
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows = _parse_rc_rows(soup, base_url)

        if not rows:
            log.info("[Regione Calabria] No rows on page %d – stopping", page)
            break

        for row in rows:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)

        if not _has_next_page(soup, page):
            break
        page += 1

    log.info("[Regione Calabria] Found %d matching acts", len(results))
    return results


def _parse_rc_rows(soup: BeautifulSoup, base_url: str) -> list[dict]:
    rows = []
    # Try table first
    table = soup.find("table")
    if table:
        headers = []
        thead = table.find("thead")
        if thead:
            headers = [th.get_text(strip=True).lower() for th in thead.find_all(["th", "td"])]
        tbody = table.find("tbody") or table
        for tr in tbody.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            if not cells or all(c.name == "th" for c in cells):
                continue
            row_data = {}
            if headers and len(headers) == len(cells):
                for h, c in zip(headers, cells):
                    row_data[h] = c.get_text(strip=True)
            else:
                texts = [c.get_text(strip=True) for c in cells]
                for i, (k, v) in enumerate(zip(
                    ["numero", "data", "tipo", "oggetto", "extra"], texts
                )):
                    row_data[k] = v

            link = None
            for a in tr.find_all("a", href=True):
                href = a["href"]
                if href and href != "#":
                    link = urljoin(base_url, href)
                    break

            oggetto = (
                row_data.get("oggetto")
                or row_data.get("titolo")
                or row_data.get("descrizione")
                or ""
            )
            data_pubbl = (
                row_data.get("data")
                or row_data.get("data pubblicazione")
                or ""
            )
            rows.append({"data": data_pubbl, "oggetto": oggetto, "link": link or ""})
        return rows

    # Fallback: list items
    items = soup.find_all(class_=re.compile(r"provvedimento|atto|item|result", re.I))
    for item in items:
        a_tag = item.find("a", href=True)
        link = urljoin(base_url, a_tag["href"]) if a_tag else ""
        oggetto = a_tag.get_text(strip=True) if a_tag else item.get_text(strip=True)
        data_el = item.find(class_=re.compile(r"data|date", re.I))
        data_pubbl = data_el.get_text(strip=True) if data_el else ""
        rows.append({"data": data_pubbl, "oggetto": oggetto, "link": link})

    return rows


# ── TAR Catanzaro ─────────────────────────────────────────────────────────────
TAR_BASE = "https://www.giustizia-amministrativa.it"
TAR_PATH = "/provvedimenti-tar-catanzaro"


def fetch_tar_catanzaro() -> list[dict]:
    results = []
    date_from_str = DATE_FROM_3D.strftime("%Y-%m-%d")
    base_url = TAR_BASE + TAR_PATH

    log.info("[TAR Catanzaro] Starting fetch from %s", base_url)

    page = 1
    while True:
        log.info("[TAR Catanzaro] Fetching page %d", page)
        params = {
            "publishDateFrom": date_from_str,
            "page": str(page),
        }
        r = safe_get(base_url, params=params)
        if r is None:
            log.error("[TAR Catanzaro] Failed to fetch page %d", page)
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows = _parse_tar_rows(soup, base_url)

        if not rows:
            log.info("[TAR Catanzaro] No rows on page %d – stopping", page)
            break

        for row in rows:
            if matches_keywords(row.get("parte", "") + " " + row.get("oggetto", "")):
                results.append(row)

        if not _has_next_page(soup, page):
            break
        page += 1

    log.info("[TAR Catanzaro] Found %d matching acts", len(results))
    return results


def _parse_tar_rows(soup: BeautifulSoup, base_url: str) -> list[dict]:
    rows = []
    table = soup.find("table")
    if not table:
        return rows

    headers = []
    thead = table.find("thead")
    if thead:
        headers = [th.get_text(strip=True).lower() for th in thead.find_all(["th", "td"])]

    tbody = table.find("tbody") or table
    for tr in tbody.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if not cells or all(c.name == "th" for c in cells):
            continue

        row_data = {}
        if headers and len(headers) == len(cells):
            for h, c in zip(headers, cells):
                row_data[h] = c.get_text(strip=True)
        else:
            texts = [c.get_text(strip=True) for c in cells]
            for k, v in zip(["data", "numero", "tipo", "parte", "oggetto", "extra"], texts):
                row_data[k] = v

        link = None
        for a in tr.find_all("a", href=True):
            href = a["href"]
            if href and href != "#":
                link = urljoin(base_url, href)
                break

        parte = (
            row_data.get("parte")
            or row_data.get("ricorrente")
            or ""
        )
        oggetto = (
            row_data.get("oggetto")
            or row_data.get("descrizione")
            or ""
        )
        data_pubbl = row_data.get("data") or row_data.get("data pubblicazione") or ""

        rows.append({
            "data": data_pubbl,
            "parte": parte,
            "oggetto": oggetto,
            "link": link or "",
        })

    return rows


# ── Excel output ──────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
LINK_FONT = Font(color="0563C1", underline="single")
ALT_FILL = PatternFill("solid", fgColor="DCE6F1")

COLUMNS = ["Data", "Oggetto", "Descrizione breve", "Link"]
COL_WIDTHS = [15, 60, 55, 12]


def build_excel(data_by_sheet: dict[str, list[dict]], out_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    sheet_order = [
        "ASP Cosenza",
        "ASP Catanzaro",
        "ASP Crotone",
        "ASP Reggio Calabria",
        "ASP Vibo Valentia",
        "Regione Calabria",
        "TAR Catanzaro",
    ]

    for sheet_name in sheet_order:
        rows = data_by_sheet.get(sheet_name, [])
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel sheet name max 31 chars

        # Header
        for col_idx, (col_name, width) in enumerate(zip(COLUMNS, COL_WIDTHS), start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        ws.row_dimensions[1].height = 20

        if not rows:
            ws.cell(row=2, column=1, value="Nessun risultato")
            continue

        for row_idx, row in enumerate(rows, start=2):
            oggetto = row.get("oggetto") or row.get("parte") or ""
            data_val = row.get("data", "")
            link_url = row.get("link", "")
            descrizione = oggetto[:150] if oggetto else ""

            # Alternate row shading
            fill = ALT_FILL if row_idx % 2 == 0 else None

            # Data
            c_data = ws.cell(row=row_idx, column=1, value=data_val)
            c_data.alignment = Alignment(vertical="top")
            if fill:
                c_data.fill = fill

            # Oggetto
            c_ogg = ws.cell(row=row_idx, column=2, value=oggetto)
            c_ogg.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                c_ogg.fill = fill

            # Descrizione breve
            c_desc = ws.cell(row=row_idx, column=3, value=descrizione)
            c_desc.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                c_desc.fill = fill

            # Link (hyperlink cliccabile)
            c_link = ws.cell(row=row_idx, column=4, value="Apri" if link_url else "")
            if link_url:
                c_link.hyperlink = link_url
                c_link.font = LINK_FONT
            c_link.alignment = Alignment(vertical="top")
            if fill:
                c_link.fill = fill

        # Freeze header
        ws.freeze_panes = "A2"

    wb.save(out_path)
    log.info("Excel saved to %s", out_path)


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    data_by_sheet = {}

    # ASP portals
    for name, base_url in ASP_PORTALS:
        data_by_sheet[name] = fetch_asp_portal(name, base_url)

    # Regione Calabria
    data_by_sheet["Regione Calabria"] = fetch_regione_calabria()

    # TAR Catanzaro
    data_by_sheet["TAR Catanzaro"] = fetch_tar_catanzaro()

    # Build filename
    date_str = TODAY.strftime("%d-%m-%Y")
    filename = f"Monitoraggio_Atti_Calabria_{date_str}.xlsx"
    out_path = f"/home/user/NS/{filename}"

    build_excel(data_by_sheet, out_path)

    # Summary
    print("\n" + "=" * 60)
    print(f"File: {filename}")
    print("=" * 60)
    for sheet, rows in data_by_sheet.items():
        print(f"  {sheet:30s}: {len(rows)} atti trovati")
    print("=" * 60)

    return out_path


if __name__ == "__main__":
    main()
