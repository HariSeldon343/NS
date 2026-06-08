#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraper per portali ASP Calabria, Regione Calabria e TAR Catanzaro

Eseguibile sia in cloud (solo Regione Calabria raggiungibile) sia localmente.
Richiede: pip install requests beautifulsoup4 openpyxl lxml

Logica date:
  - ASP/Regione: ultimi 2 giorni (lunedì: 3 giorni per includere il venerdì)
  - TAR: ultimi 3 giorni (lunedì: 4 giorni)
"""

import re
import sys
import time
import traceback
from datetime import datetime, timedelta
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Date setup ────────────────────────────────────────────────────────────────
TODAY       = datetime.today()
IS_MONDAY   = TODAY.weekday() == 0   # 0 = Monday

# Se lunedì, includi anche il venerdì (3 giorni per ASP/Regione, 4 per TAR)
DAYS_BACK_ASP = 3 if IS_MONDAY else 2
DAYS_BACK_TAR = 4 if IS_MONDAY else 3

DATE_FROM_ASP = (TODAY - timedelta(days=DAYS_BACK_ASP)).strftime("%Y-%m-%d")
DATE_FROM_TAR = (TODAY - timedelta(days=DAYS_BACK_TAR)).strftime("%Y-%m-%d")
DATE_TO       = TODAY.strftime("%Y-%m-%d")

print(f"[INFO] Esecuzione: {TODAY.strftime('%d/%m/%Y %H:%M')} ({'lunedì' if IS_MONDAY else TODAY.strftime('%A')})")
print(f"[INFO] Filtro ASP/Regione: {DATE_FROM_ASP} → {DATE_TO} ({DAYS_BACK_ASP} gg)")
print(f"[INFO] Filtro TAR:         {DATE_FROM_TAR} → {DATE_TO} ({DAYS_BACK_TAR} gg)")

# ── Keywords ──────────────────────────────────────────────────────────────────
KEYWORDS_RAW = [
    r"\bADI\b",
    r"Assistenza domiciliare",
    r"\bANMIC\b",
    r"Accreditamento",
    r"Aumento di budget",
    r"Autismo",
    r"Autorizzazione all[''']esercizio",
    r"Autorizzazione alla realizzazione",
    r"\bAutorizzazioni\b",
    r"\bBudget\b",
    r"Casa Giardino",
    r"Centro San Giuseppe",
    r"Centro salute e benessere",
    r"Fabbisogni LEA",
    r"Fisiolab",
    r"Fisioterapia",
    r"\bLIFE\b",
    r"Parere commissione",
    r"Presa d[''']atto verifica",
    r"Programmazione",
    r"Rete riabilitativa",
    r"Rete territoriale",
    r"Riabilitazione estensiva",
    r"Riconversione prestazioni",
    r"Rinnovo accreditamento",
    r"San Teodoro",
    r"Savelli Hospital",
    r"Starbene",
    r"Verifica requisiti",
    r"Villa San Giuseppe",
    r"Villa del Rosario",
]
KW_PATTERNS = [re.compile(kw, re.IGNORECASE) for kw in KEYWORDS_RAW]


def matches_keywords(text: str) -> bool:
    for pat in KW_PATTERNS:
        if pat.search(text):
            return True
    return False


# ── HTTP session ──────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# ── ASP AlboOnline scraper ────────────────────────────────────────────────────

ASP_PORTALS = {
    "ASP Cosenza":         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def scrape_asp_portal(name: str, base_url: str) -> list[dict]:
    """Scrape a single ASP AlboOnline portal."""
    print(f"\n{'='*60}\n[ASP] {name}")
    results: list[dict] = []
    session = make_session()
    session.headers["Referer"] = base_url

    # 1. GET page to obtain session cookies + form details
    r0 = session.get(base_url, timeout=30, verify=False)
    r0.raise_for_status()
    soup0 = BeautifulSoup(r0.text, "lxml")
    print(f"  Title: {(soup0.title.string or '').strip()}")

    form = soup0.find("form")
    if not form:
        print("  [WARN] No form found on page")
        return []

    form_action = form.get("action") or base_url
    if not form_action.startswith("http"):
        form_action = urljoin(base_url, form_action)
    form_method = form.get("method", "get").upper()
    print(f"  Form: {form_method} {form_action}")

    # Collect all hidden + existing values
    payload: dict[str, str] = {}
    for inp in form.find_all("input"):
        n, v = inp.get("name", ""), inp.get("value", "")
        if n:
            payload[n] = v
    for sel in form.find_all("select"):
        n = sel.get("name", "")
        opt = sel.find("option", selected=True)
        if n:
            payload[n] = opt["value"] if opt and opt.get("value") else ""

    # Set date filter
    date_set = False
    for inp in form.find_all("input"):
        iid, iname = inp.get("id", ""), inp.get("name", "")
        if "dataPubblicazioneDal" in iid or "dataPubblicazioneDal" in iname:
            key = iname or iid
            payload[key] = DATE_FROM_ASP
            print(f"  Date field '{key}' = {DATE_FROM_ASP}")
            date_set = True
    if not date_set:
        for cand in ["dataPubblicazioneDal", "form:dataPubblicazioneDal", "dataDal", "dataInizio"]:
            payload[cand] = DATE_FROM_ASP

    # 2. Paginate
    page = 1
    seen_sigs: set[str] = set()
    while True:
        print(f"  Page {page}... ", end="")
        if form_method == "POST":
            resp = session.post(form_action, data=payload, timeout=30, verify=False)
        else:
            resp = session.get(form_action, params=payload, timeout=30, verify=False)

        if resp.status_code != 200:
            print(f"HTTP {resp.status_code}")
            break

        page_soup = BeautifulSoup(resp.text, "lxml")
        rows = _parse_asp_table(page_soup, base_url)
        print(f"{len(rows)} rows")

        if not rows:
            break
        sig = "|".join(r["oggetto"][:30] for r in rows)
        if sig in seen_sigs:
            print("  [STOP] Duplicate page, stopping")
            break
        seen_sigs.add(sig)

        for row in rows:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)

        # Check for next page
        nxt = _asp_next_payload(page_soup, payload, page)
        if not nxt:
            break
        payload = nxt
        page += 1
        time.sleep(0.8)

    print(f"  => {len(results)} matching")
    return results


def _parse_asp_table(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Extract rows from an ASP AlboOnline result table."""
    # Find the most relevant table
    target = None
    for table in soup.find_all("table"):
        header_text = " ".join(
            th.get_text(strip=True).lower()
            for th in (table.find("thead") or table).find_all("th")
        )
        if any(k in header_text for k in ("oggetto", "descrizione", "atto", "numero")):
            target = table
            break
    if not target:
        all_tables = soup.find_all("table")
        target = max(all_tables, key=lambda t: len(t.find_all("tr"))) if all_tables else None
    if not target:
        return []

    # Get header column names
    thead = target.find("thead")
    header_row = (thead or target).find("tr")
    headers = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])] if header_row else []

    results = []
    # Data rows: skip first row if no thead (= header row)
    tbody = target.find("tbody")
    data_rows = tbody.find_all("tr") if tbody else target.find_all("tr")[1:]

    for tr in data_rows:
        # Use ALL cells (td + th) to match header count
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        texts = [c.get_text(strip=True) for c in cells]
        col = dict(zip(headers, texts)) if headers else {}

        link = ""
        for a in tr.find_all("a", href=True):
            href = a["href"]
            if href.startswith("#") or href.startswith("javascript"):
                continue
            link = href if href.startswith("http") else urljoin(base_url, href)
            break

        # Map to standard fields
        oggetto = (
            col.get("Oggetto") or col.get("oggetto") or
            col.get("Descrizione") or col.get("Titolo") or
            (texts[4] if len(texts) > 4 else texts[2] if len(texts) > 2 else texts[-1] if texts else "")
        )
        data_pub = (
            col.get("Data Pubblicazione") or col.get("Data") or col.get("data") or
            (texts[1] if len(texts) > 1 else texts[0] if texts else "")
        )

        if oggetto and oggetto != "→" and oggetto != ">":
            results.append({"data": data_pub, "oggetto": oggetto, "link": link})

    return results


def _asp_next_payload(soup: BeautifulSoup, cur: dict, page: int) -> dict | None:
    has_next = bool(soup.find("a", string=re.compile(r"Successiv|»|›|Next", re.I)))
    if not has_next:
        return None
    nxt = dict(cur)
    for pname in ["page", "pagina", "pageNum", "start"]:
        if pname in nxt:
            nxt[pname] = str(page)  # next page index
            return nxt
    nxt["pagina"] = str(page + 1)
    return nxt


# ── Regione Calabria scraper ──────────────────────────────────────────────────

REGIONE_URL = "https://www.regione.calabria.it/provvedimenti-della-regione"


def scrape_regione_calabria() -> list[dict]:
    """
    Scrape Provvedimenti della Regione Calabria.
    Table structure: thead with th headers, tbody rows with td+th(N) cells.
    Date filter via GET params: filter_date_from, filter_date_to, filter_active.
    Pagination via pageNum=0,1,2,...
    """
    print(f"\n{'='*60}\n[REGIONE] Regione Calabria")
    results: list[dict] = []
    session = make_session()
    session.headers["Referer"] = "https://www.regione.calabria.it/"

    page_num = 0
    seen_sigs: set[str] = set()

    while True:
        params: dict[str, str] = {
            "filter_date_from": DATE_FROM_ASP,
            "filter_date_to":   DATE_TO,
            "filter_active":    "true",
        }
        if page_num > 0:
            params["pageNum"] = str(page_num)

        print(f"  pageNum={page_num}... ", end="")
        r = session.get(REGIONE_URL, params=params, timeout=30, verify=False)
        print(f"HTTP {r.status_code}")
        if r.status_code != 200:
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows = _parse_regione_table(soup)
        print(f"  Rows found: {len(rows)}")

        if not rows:
            break
        sig = "|".join(row["oggetto"][:30] for row in rows)
        if sig in seen_sigs:
            print("  [STOP] Same content, stopping")
            break
        seen_sigs.add(sig)

        for row in rows:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)

        if not _regione_has_next(soup):
            break
        page_num += 1
        time.sleep(0.5)

    print(f"  => {len(results)} matching")
    return results


def _parse_regione_table(soup: BeautifulSoup) -> list[dict]:
    """
    Parse Regione Calabria provvedimenti table.
    Headers in <thead>, data in <tbody> with structure:
      td(Tipologia) | td(Data) | th(N) | td(Dipartimento) | td(Oggetto) | td(link→)
    """
    table = soup.find("table")
    if not table:
        return []

    # Get column headers from thead
    thead = table.find("thead")
    if not thead:
        return []
    header_tr = thead.find("tr")
    if not header_tr:
        return []
    headers = [th.get_text(strip=True) for th in header_tr.find_all(["th", "td"])]
    # Expected: ['Tipologia', 'Data Repertoriazione', 'N', 'Dipartimento', 'Oggetto', 'Dettaglio']

    tbody = table.find("tbody")
    if not tbody:
        return []

    rows: list[dict] = []
    for tr in tbody.find_all("tr"):
        # Get ALL cells (td + th) to include the decree number th
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        texts = [c.get_text(strip=True) for c in cells]

        # Map by header position (zip stops at shortest)
        col = dict(zip(headers, texts))

        # Extract hyperlink from last cell (Dettaglio → )
        link = ""
        for a in tr.find_all("a", href=True):
            href = a["href"]
            link = href if href.startswith("http") else urljoin(REGIONE_URL, href)
            break

        # Column mapping
        oggetto  = col.get("Oggetto", "").strip()
        data_pub = col.get("Data Repertoriazione", "").strip() or col.get("Data", "").strip()
        tipologia = col.get("Tipologia", "").strip()
        numero    = col.get("N", "").strip()
        dip       = col.get("Dipartimento", "").strip()

        # Fallback by position (Tipologia|Data|N|Dip|Oggetto|Dettaglio = cols 0-5)
        if not oggetto and len(texts) >= 5:
            oggetto  = texts[4].strip()
            data_pub = texts[1].strip()

        if oggetto and oggetto not in ("→", ">", "Dettaglio"):
            rows.append({
                "data":         data_pub,
                "oggetto":      oggetto,
                "tipologia":    tipologia,
                "numero":       numero,
                "dipartimento": dip,
                "link":         link,
            })

    return rows


def _regione_has_next(soup: BeautifulSoup) -> bool:
    """Check if there's a 'next page' link below the results table."""
    # The page has a pagination widget below the table
    # Look for any <a> with aria or text indicating 'successivo'
    # Exclude the main navbar
    for a in soup.find_all("a", href=True):
        txt = a.get_text(strip=True).lower()
        classes = " ".join(a.get("class", []))
        # Skip navbar links
        if "navbar" in classes or "nav-link" in classes:
            continue
        inner_spans = [s.get_text(strip=True).lower() for s in a.find_all("span")]
        all_text = txt + " " + " ".join(inner_spans)
        if any(k in all_text for k in ("successiv", "next", "›", "»")):
            return True
    return False


# ── TAR Catanzaro scraper ─────────────────────────────────────────────────────

TAR_BASE = "https://www.giustizia-amministrativa.it"
TAR_URL  = f"{TAR_BASE}/provvedimenti-tar-catanzaro"


def scrape_tar_catanzaro() -> list[dict]:
    """Scrape TAR Catanzaro provvedimenti (Liferay portal)."""
    print(f"\n{'='*60}\n[TAR] TAR Catanzaro")
    results: list[dict] = []
    session = make_session()
    session.headers["Referer"] = TAR_BASE + "/"

    page = 1
    delta = 20
    seen_sigs: set[str] = set()

    while True:
        params = {
            "publishDateFrom": DATE_FROM_TAR,
            "publishDateTo":   DATE_TO,
            "_ProvvedimentiGiustizia_WAR_ProvvedimentiGiustiziaportlet_publishDateFrom": DATE_FROM_TAR,
            "_ProvvedimentiGiustizia_WAR_ProvvedimentiGiustiziaportlet_publishDateTo":   DATE_TO,
            "_ProvvedimentiGiustizia_WAR_ProvvedimentiGiustiziaportlet_cur":   str(page),
            "_ProvvedimentiGiustizia_WAR_ProvvedimentiGiustiziaportlet_delta": str(delta),
        }
        print(f"  Page {page}... ", end="")
        r = session.get(TAR_URL, params=params, timeout=30, verify=False)
        print(f"HTTP {r.status_code}")
        r.raise_for_status()
        if r.status_code != 200:
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows = _parse_tar_table(soup)
        print(f"  Rows: {len(rows)}")
        if not rows:
            break

        sig = "|".join(r["oggetto"][:30] for r in rows)
        if sig in seen_sigs:
            break
        seen_sigs.add(sig)

        for row in rows:
            combined = (row.get("parte", "") + " " + row.get("oggetto", "")).strip()
            if matches_keywords(combined):
                results.append(row)

        has_next = bool(soup.find("a", string=re.compile(r"Successiv|›|»|Next", re.I)))
        if not has_next or len(rows) < delta:
            break
        page += 1
        time.sleep(0.8)

    print(f"  => {len(results)} matching")
    return results


def _parse_tar_table(soup: BeautifulSoup) -> list[dict]:
    rows: list[dict] = []
    for table in soup.find_all("table"):
        ths = [th.get_text(strip=True).lower() for th in table.find_all("th")]
        if any(k in " ".join(ths) for k in ("parte", "numero", "tipo", "data")):
            headers = [th.get_text(strip=True) for th in table.find_all("th")]
            tbody = table.find("tbody")
            trs = tbody.find_all("tr") if tbody else table.find_all("tr")[1:]
            for tr in trs:
                cells = tr.find_all(["td", "th"])
                if not cells:
                    continue
                texts = [c.get_text(strip=True) for c in cells]
                col = dict(zip(headers, texts))
                link = ""
                for a in tr.find_all("a", href=True):
                    href = a["href"]
                    link = href if href.startswith("http") else urljoin(TAR_BASE, href)
                    break
                parte   = col.get("Parte", "") or col.get("parte", "")
                oggetto = col.get("Oggetto", "") or col.get("Tipo", "") or col.get("tipo", "")
                data_p  = col.get("Data", "") or col.get("Data Pubblicazione", "")
                display = f"{parte} | {oggetto}".strip(" |") if parte else oggetto
                rows.append({
                    "data":   data_p,
                    "parte":  parte,
                    "oggetto": display,
                    "link":   link,
                })
            break
    return rows


# ── Excel output ──────────────────────────────────────────────────────────────

SHEET_ORDER = [
    "ASP Cosenza",
    "ASP Catanzaro",
    "ASP Crotone",
    "ASP Reggio Calabria",
    "ASP Vibo Valentia",
    "Regione Calabria",
    "TAR Catanzaro",
]

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
LNK_FONT = Font(color="0563C1", underline="single", name="Calibri", size=10)
BDY_FONT = Font(name="Calibri", size=10)
ALT_FILL = PatternFill("solid", fgColor="DCE6F1")
WRN_FONT = Font(color="C00000", italic=True, name="Calibri", size=10)


def write_excel(all_results: dict, output_path: str, notes: dict | None = None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_name in SHEET_ORDER:
        rows = all_results.get(sheet_name, [])
        note = (notes or {}).get(sheet_name, "")
        ws   = wb.create_sheet(title=sheet_name[:31])

        # Header
        for ci, h in enumerate(["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"], 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font      = HDR_FONT
            cell.fill      = HDR_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        # Data or notice
        if note and not rows:
            nc = ws.cell(row=2, column=1, value=note)
            nc.font = WRN_FONT
            nc.alignment = Alignment(horizontal="left", wrap_text=True)
            ws.merge_cells("A2:D2")
            ws.row_dimensions[2].height = 40
        elif not rows:
            nc = ws.cell(row=2, column=1, value="Nessun risultato nel periodo selezionato")
            nc.alignment = Alignment(horizontal="center")
            nc.font = BDY_FONT
            ws.merge_cells("A2:D2")
        else:
            for ri, row in enumerate(rows, 2):
                oggetto  = row.get("oggetto", "")
                data_val = row.get("data", "")
                link     = row.get("link", "").strip()
                desc     = oggetto[:150]
                fill     = ALT_FILL if ri % 2 == 0 else PatternFill()

                def _cell(col_i, val, font=BDY_FONT, wrap=False):
                    c = ws.cell(row=ri, column=col_i, value=val)
                    c.font = font
                    c.fill = fill
                    c.alignment = Alignment(vertical="top", wrap_text=wrap)
                    return c

                _cell(1, data_val)
                _cell(2, oggetto, wrap=True)
                _cell(3, desc,    wrap=True)
                lc = _cell(4, link or "N/D")
                if link:
                    lc.hyperlink = link
                    lc.font = LNK_FONT

        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 72
        ws.column_dimensions["C"].width = 42
        ws.column_dimensions["D"].width = 52
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:D{max(2, len(rows) + 1)}"

    wb.save(output_path)
    print(f"\n[OK] Salvato: {output_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def run_scraping():
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    all_results: dict[str, list[dict]] = {}
    notes:       dict[str, str]        = {}

    # ASP portals
    for pname, purl in ASP_PORTALS.items():
        try:
            all_results[pname] = scrape_asp_portal(pname, purl)
        except (requests.HTTPError, requests.ConnectionError) as e:
            code = getattr(getattr(e, "response", None), "status_code", None)
            if code == 503 or "503" in str(e) or "timeout" in str(e).lower() or "upstream" in str(e).lower():
                notes[pname] = (
                    f"⚠ Portale non raggiungibile da questa rete (errore 503 / IP filtering del proxy cloud). "
                    f"Eseguire lo script localmente per ottenere i dati. URL: {purl}"
                )
            else:
                notes[pname] = f"Errore di rete: {e}"
            all_results[pname] = []
            print(f"  [SKIP] {pname}: {notes[pname][:80]}...")
        except Exception as e:
            notes[pname] = f"Errore imprevisto: {e}"
            all_results[pname] = []

    # Regione Calabria
    try:
        all_results["Regione Calabria"] = scrape_regione_calabria()
    except Exception as e:
        notes["Regione Calabria"] = f"Errore: {e}"
        all_results["Regione Calabria"] = []

    # TAR Catanzaro
    try:
        all_results["TAR Catanzaro"] = scrape_tar_catanzaro()
    except (requests.HTTPError, requests.ConnectionError) as e:
        code = getattr(getattr(e, "response", None), "status_code", None)
        if code == 503 or "503" in str(e):
            notes["TAR Catanzaro"] = (
                "⚠ Portale TAR non raggiungibile da questa rete (503 - IP filtering). "
                "Eseguire lo script localmente."
            )
        else:
            notes["TAR Catanzaro"] = f"Errore: {e}"
        all_results["TAR Catanzaro"] = []
    except Exception as e:
        notes["TAR Catanzaro"] = f"Errore: {e}"
        all_results["TAR Catanzaro"] = []

    return all_results, notes


def main() -> str:
    all_results, notes = run_scraping()

    print(f"\n{'='*60}\nRIEPILOGO:")
    total = 0
    for sheet in SHEET_ORDER:
        n = len(all_results.get(sheet, []))
        total += n
        tag = " [NON ACCESSIBILE]" if notes.get(sheet) else ""
        print(f"  {sheet:25s}: {n:3d} atti matchati{tag}")
    print(f"  {'TOTALE':25s}: {total}")

    date_str = TODAY.strftime("%d-%m-%Y")
    out_path = f"Monitoraggio_Atti_Calabria_{date_str}.xlsx"
    write_excel(all_results, out_path, notes)
    return out_path


if __name__ == "__main__":
    out = main()
    print(f"\n→ {out}")
