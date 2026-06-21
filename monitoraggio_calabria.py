#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scrape: 5 ASP portals, Regione Calabria, TAR Catanzaro → XLSX report.
"""

import re, time, datetime, sys
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Dates ────────────────────────────────────────────────────────────────────
TODAY        = datetime.date.today()
DATE_FROM_2  = TODAY - datetime.timedelta(days=2)   # ASP + Regione: last 2 days
DATE_FROM_3  = TODAY - datetime.timedelta(days=3)   # TAR: last 3 days

print(f"TODAY={TODAY}  |  cutoff-2d={DATE_FROM_2}  |  cutoff-3d={DATE_FROM_3}")

# ── Keywords ─────────────────────────────────────────────────────────────────
KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo",
    "Autorizzazione all'esercizio", "Autorizzazione alla realizzazione",
    "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento",
    "San Teodoro", "Savelli Hospital", "Starbene",
    "Verifica requisiti", "Villa San Giuseppe", "Villa del Rosario",
]
LIFE_RE     = re.compile(r'\bLIFE\b', re.IGNORECASE)
KW_PATTERNS = [re.compile(re.escape(kw), re.IGNORECASE) for kw in KEYWORDS]


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    if LIFE_RE.search(text):
        return True
    return any(p.search(text) for p in KW_PATTERNS)


# ── HTTP session ──────────────────────────────────────────────────────────────
def make_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/125.0.0.0 Safari/537.36",
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection":      "keep-alive",
    })
    return s


def get_page(session, url, params=None, retries=2, wait=2):
    for attempt in range(retries):
        try:
            r = session.get(url, params=params, timeout=30)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  [GET retry {attempt+1}] {url[:80]} -> {e}")
            time.sleep(wait * (attempt + 1))
    return None


def post_page(session, url, data, retries=2, wait=2):
    for attempt in range(retries):
        try:
            r = session.post(url, data=data, timeout=30)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  [POST retry {attempt+1}] {url[:80]} -> {e}")
            time.sleep(wait * (attempt + 1))
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER 1 – ASP portals (sisr.regione.calabria.it)
# ═══════════════════════════════════════════════════════════════════════════════

ASP_PORTALS = {
    "ASP Cosenza":         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def scrape_asp(name: str, base_url: str) -> list[dict]:
    print(f"\n=== {name} ===")
    session = make_session()
    results = []

    # Phase 1: GET the form page to capture hidden fields + form action
    r0 = get_page(session, base_url)
    if r0 is None:
        print(f"  Portale non raggiungibile (503/timeout)")
        return results

    soup0 = BeautifulSoup(r0.text, "lxml")
    form  = soup0.find("form")
    if not form:
        print(f"  Nessun form trovato")
        return results

    hidden = {inp["name"]: inp.get("value", "")
              for inp in form.find_all("input", type="hidden")
              if inp.get("name")}

    action = form.get("action", base_url)
    if not action.startswith("http"):
        from urllib.parse import urlparse, urljoin
        action = urljoin(base_url, action)

    # Phase 2: iterate pages
    page = 1
    while True:
        post_data = {
            **hidden,
            "dataPubblicazioneDal": DATE_FROM_2.isoformat(),
            "dataPubblicazioneAl":  TODAY.isoformat(),
        }
        if page > 1:
            for k in ("page", "pagina", "currentPage", "pg"):
                post_data[k] = page

        r = post_page(session, action, post_data)
        if r is None:
            print(f"  Errore pagina {page}")
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows, has_more = _parse_asp_page(soup, base_url, page)

        found = 0
        for row in rows:
            if matches_keywords(row.get("oggetto", "")):
                results.append(row)
                found += 1
        print(f"  Pag {page}: {len(rows)} righe, {found} match")

        if not rows or not has_more or page >= 50:
            break
        page += 1
        time.sleep(0.5)

    return results


def _parse_asp_page(soup, base_url, page_num):
    """Return (rows, has_more)."""
    rows = []
    from urllib.parse import urljoin

    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if len(trs) < 2:
            continue
        hdrs = [th.get_text(strip=True).lower()
                for th in trs[0].find_all(["th", "td"])]
        # must look like a results table
        if not any(h for h in hdrs if "oggetto" in h or "titolo" in h or "atto" in h or "data" in h):
            continue
        for tr in trs[1:]:
            cells = tr.find_all("td")
            if not cells:
                continue
            row = {}
            for i, cell in enumerate(cells):
                txt  = cell.get_text(strip=True)
                href = ""
                a    = cell.find("a")
                if a and a.get("href"):
                    href = a["href"]
                    if not href.startswith("http"):
                        href = urljoin(base_url, href)

                lbl = hdrs[i] if i < len(hdrs) else ""
                if "data" in lbl and "oggetto" not in row:
                    row["data"] = txt
                elif "oggetto" in lbl or "titolo" in lbl:
                    row["oggetto"] = txt
                    if href:
                        row["link"] = href
                elif not row:
                    row["data"] = txt
                elif "oggetto" not in row:
                    row["oggetto"] = txt
                    if href:
                        row["link"] = href
                if href and "link" not in row:
                    row["link"] = href

            if row.get("oggetto") or row.get("data"):
                rows.append(row)

    # Pagination detection
    has_more = False
    for a in soup.find_all("a"):
        txt  = a.get_text(strip=True).lower()
        href = a.get("href", "")
        if any(kw in txt for kw in ["successiv", "next", "»", ">"]):
            has_more = True
            break
        if f"page={page_num + 1}" in href or f"pagina={page_num + 1}" in href:
            has_more = True
            break
        try:
            if int(a.get_text(strip=True)) > page_num:
                has_more = True
                break
        except (ValueError, TypeError):
            pass

    return rows, has_more


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER 2 – Regione Calabria
# ═══════════════════════════════════════════════════════════════════════════════

REGIONE_URL = "https://www.regione.calabria.it/provvedimenti-della-regione/"


def scrape_regione() -> list[dict]:
    print("\n=== Regione Calabria ===")
    session = make_session()
    results = []
    page    = 1
    done    = False

    while not done and page <= 100:
        params = {} if page == 1 else {"paged": page}
        r = get_page(session, REGIONE_URL, params=params)
        if r is None:
            print("  Portale non raggiungibile")
            break

        soup  = BeautifulSoup(r.text, "lxml")
        table = soup.find("table")
        if not table:
            print(f"  Pag {page}: nessuna tabella – stop")
            break

        trs = table.find_all("tr")[1:]   # skip header
        if not trs:
            print(f"  Pag {page}: nessuna riga – stop")
            break

        page_rows = []
        for tr in trs:
            cells = tr.find_all(["td", "th"])
            if len(cells) < 5:
                continue

            data_str = cells[1].get_text(strip=True)
            oggetto  = cells[4].get_text(strip=True)
            a        = cells[-1].find("a")
            link     = a["href"] if a and a.get("href") else ""

            try:
                d = datetime.datetime.strptime(data_str, "%d/%m/%Y").date()
            except ValueError:
                d = None

            if d and d < DATE_FROM_2:
                done = True      # rows are in descending order → no more in range
                break

            page_rows.append({"data": data_str, "oggetto": oggetto, "link": link})

        found = 0
        for row in page_rows:
            if matches_keywords(row["oggetto"]):
                results.append(row)
                found += 1
        print(f"  Pag {page}: {len(page_rows)} righe nel range, {found} match")

        if done:
            break

        # Check pagination
        has_more = any(
            a.get_text(strip=True).lower() in ("»", "›", "next", "successiv")
            or f"paged={page + 1}" in a.get("href", "")
            for a in soup.find_all("a")
        )
        if not has_more:
            break
        page += 1
        time.sleep(0.3)

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER 3 – TAR Catanzaro (Liferay portlet)
# ═══════════════════════════════════════════════════════════════════════════════

TAR_URL = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"


def scrape_tar() -> list[dict]:
    print("\n=== TAR Catanzaro ===")
    session = make_session()
    results = []

    # GET base page to extract dynamic tokens
    r0 = get_page(session, TAR_URL)
    if r0 is None:
        print("  Portale non raggiungibile")
        return results

    soup0 = BeautifulSoup(r0.text, "lxml")

    # Find the portlet search form
    form = None
    for f in soup0.find_all("form"):
        action = f.get("action", "")
        if "JurisdictionalActivity" in action and "lifecycle=1" in action:
            form = f
            break

    if not form:
        print("  Form portlet non trovato")
        return results

    form_action = form["action"]
    m_pid = re.search(r'p_p_id=([^&]+)&', form_action)
    if not m_pid:
        print("  portlet ID non trovato nel form action")
        return results

    portlet_id = m_pid.group(1)
    PREFIX     = f"_{portlet_id}_"

    form_date_inp = form.find("input", {"name": f"{PREFIX}formDate"})
    form_date_val = form_date_inp["value"] if form_date_inp else str(int(time.time() * 1000))

    session.headers.update({
        "Referer": TAR_URL,
        "Origin":  "https://www.giustizia-amministrativa.it",
    })

    page  = 0        # 0-based Liferay delta
    delta = 20

    while True:
        post_data = {
            f"{PREFIX}formDate":        form_date_val,
            f"{PREFIX}year":            "",
            f"{PREFIX}number":          "",
            f"{PREFIX}hearingDateFrom": "",
            f"{PREFIX}hearingDateTo":   "",
            f"{PREFIX}section":         "",
            f"{PREFIX}type":            "",
            f"{PREFIX}specific":        "",
            f"{PREFIX}publishDateFrom": DATE_FROM_3.isoformat(),
            f"{PREFIX}publishDateTo":   TODAY.isoformat(),
            f"{PREFIX}president":       "",
            f"{PREFIX}draftingJudge":   "",
        }
        # Liferay pagination
        if page > 0:
            post_data[f"{PREFIX}cur"]   = page + 1
            post_data[f"{PREFIX}delta"] = delta

        r1 = post_page(session, form_action, post_data)
        if r1 is None:
            print(f"  Errore POST pag {page+1}")
            break

        soup1 = BeautifulSoup(r1.text, "lxml")

        # Check for error message
        err_elem = soup1.find(class_=re.compile(r"portlet-msg-error|alert-danger|error"))
        if err_elem:
            print(f"  Errore server: {err_elem.get_text(strip=True)[:120]}")
            break

        rows = _parse_tar_page(soup1)
        if not rows:
            print(f"  Pag {page+1}: nessuna riga – stop")
            break

        found = 0
        for row in rows:
            parte = row.get("parte", "") or row.get("oggetto", "")
            if matches_keywords(parte):
                results.append(row)
                found += 1
        print(f"  Pag {page+1}: {len(rows)} righe, {found} match")

        # Pagination
        if not _tar_has_next(soup1, page + 1) or page >= 50:
            break
        page += 1
        time.sleep(0.5)

    return results


def _parse_tar_page(soup) -> list[dict]:
    rows = []
    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if len(trs) < 2:
            continue
        hdrs = [th.get_text(strip=True).lower()
                for th in trs[0].find_all(["th", "td"])]
        if not any("parte" in h or "nrg" in h or "numero" in h for h in hdrs):
            continue
        for tr in trs[1:]:
            cells = tr.find_all("td")
            if not cells:
                continue
            row = {}
            a_tag = tr.find("a")
            if a_tag and a_tag.get("href"):
                href = a_tag["href"]
                if not href.startswith("http"):
                    href = "https://www.giustizia-amministrativa.it" + href
                row["link"] = href
            for i, cell in enumerate(cells):
                txt = cell.get_text(strip=True)
                lbl = hdrs[i] if i < len(hdrs) else ""
                if "nrg" in lbl or "numero" in lbl:
                    row["numero"] = txt
                elif "sez" in lbl:
                    row["sezione"] = txt
                elif "parte" in lbl:
                    row["parte"] = txt
                elif "tipo udienza" in lbl:
                    row["tipo_udienza"] = txt
                elif "data udienza" in lbl:
                    row["data_udienza"] = txt
                elif "numero provvedimento" in lbl:
                    row["num_provv"] = txt
                elif "data pubblicazione" in lbl or "pubbl" in lbl:
                    row["data"] = txt
                elif "tipo provvedimento" in lbl:
                    row["tipo_provv"] = txt
                elif "relatore" in lbl:
                    row["relatore"] = txt
                elif "presidente" in lbl:
                    row["presidente"] = txt
                elif "esito" in lbl:
                    row["esito"] = txt
            # Build display oggetto
            row["oggetto"] = " | ".join(filter(None, [
                row.get("numero"),
                row.get("sezione"),
                row.get("parte"),
                row.get("tipo_provv"),
                row.get("data"),
            ]))
            if row.get("parte") or row.get("numero"):
                rows.append(row)
    return rows


def _tar_has_next(soup, current_page):
    for a in soup.find_all("a"):
        txt  = a.get_text(strip=True).lower()
        href = a.get("href", "")
        if any(k in txt for k in ["successiv", "next", "»", ">"]):
            return True
        if f"cur={current_page + 1}" in href:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════════════════════════════════════════════

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
LINK_FONT   = Font(color="0563C1", underline="single", size=10)
NORMAL_FONT = Font(size=10)
ALT_FILL    = PatternFill("solid", fgColor="D9E8F5")
COLUMNS     = ["Data", "Oggetto", "Descrizione breve", "Link"]
COL_WIDTHS  = [14, 65, 45, 40]


def write_sheet(ws, rows: list[dict]):
    # Header row
    for col, (title, width) in enumerate(zip(COLUMNS, COL_WIDTHS), 1):
        cell            = ws.cell(row=1, column=col, value=title)
        cell.fill       = HEADER_FILL
        cell.font       = HEADER_FONT
        cell.alignment  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    if not rows:
        cell      = ws.cell(row=2, column=1, value="Nessun risultato")
        cell.font = Font(italic=True, color="888888", size=10)
        return

    for r_idx, row in enumerate(rows, 2):
        fill_bg = ALT_FILL if r_idx % 2 == 0 else None

        data_val = row.get("data", "")
        oggetto  = row.get("oggetto", "") or row.get("parte", "")
        descr    = oggetto[:150]
        link_url = row.get("link", "")

        for col, val in enumerate([data_val, oggetto, descr, link_url], 1):
            cell           = ws.cell(row=r_idx, column=col, value=val)
            cell.font      = NORMAL_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=(col == 2))
            if fill_bg:
                cell.fill = fill_bg

        # Clickable hyperlink in column 4
        if link_url:
            lc           = ws.cell(row=r_idx, column=4)
            lc.value     = link_url
            lc.hyperlink = link_url
            lc.font      = LINK_FONT


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

SHEET_ORDER = [
    "ASP Cosenza",
    "ASP Catanzaro",
    "ASP Crotone",
    "ASP Reggio Calabria",
    "ASP Vibo Valentia",
    "Regione Calabria",
    "TAR Catanzaro",
]


def main():
    data = {}

    for name, url in ASP_PORTALS.items():
        data[name] = scrape_asp(name, url)

    data["Regione Calabria"] = scrape_regione()
    data["TAR Catanzaro"]    = scrape_tar()

    print("\n" + "=" * 60)
    print("RIEPILOGO")
    print("=" * 60)
    total = 0
    for name in SHEET_ORDER:
        n = len(data.get(name, []))
        total += n
        print(f"  {name}: {n}")
    print(f"  TOTALE match: {total}")

    # Build workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for name in SHEET_ORDER:
        ws = wb.create_sheet(title=name)
        write_sheet(ws, data.get(name, []))
        print(f"  Foglio '{name}': OK")

    fname    = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    out_path = f"/home/user/NS/{fname}"
    wb.save(out_path)
    print(f"\nFile salvato: {out_path}")
    return out_path, total


if __name__ == "__main__":
    out, tot = main()
    print(f"\nDone. Totale match: {tot}")
