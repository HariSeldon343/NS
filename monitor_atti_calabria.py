#!/usr/bin/env python3
"""
Monitor Atti Calabria
Scraper per ASP portals, Regione Calabria e TAR Catanzaro.
"""

import re
import time
import warnings
import base64
from datetime import date, timedelta
from io import BytesIO

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ── Date range ─────────────────────────────────────────────────────────────────
TODAY = date.today()
DATE_FROM_2D = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM_3D = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")

# ── Keywords ───────────────────────────────────────────────────────────────────
KEYWORDS_SIMPLE = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo",
    "Autorizzazione all'esercizio", "Autorizzazione alla realizzazione",
    "Autorizzazioni", "Budget", "Casa Giardino", "Centro San Giuseppe",
    "Centro salute e benessere", "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]
RE_SIMPLE = re.compile(
    "|".join(re.escape(k) for k in KEYWORDS_SIMPLE), re.IGNORECASE
)
RE_LIFE = re.compile(r"\bLIFE\b", re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    return bool(RE_SIMPLE.search(text) or RE_LIFE.search(text))


# ── HTTP helpers ───────────────────────────────────────────────────────────────
BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
}


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(BROWSER_HEADERS)
    return s


# ── ASP portals ────────────────────────────────────────────────────────────────
ASP_PORTALS = {
    "ASP Cosenza":        "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":      "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":        "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria":"https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":  "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def _asp_get_hidden(soup) -> dict:
    """Extract hidden form fields."""
    hidden = {}
    for inp in soup.find_all("input", type="hidden"):
        name = inp.get("name") or inp.get("id", "")
        val  = inp.get("value", "")
        if name:
            hidden[name] = val
    return hidden


def _asp_parse_rows(soup) -> list[dict]:
    """Parse result rows from an ASP AlboOnline page."""
    rows_data = []
    table = soup.find("table", id=re.compile("alboTable|tabellaRisultati|risultati", re.I))
    if table is None:
        table = soup.find("table", class_=re.compile("table|result", re.I))
    if table is None:
        return rows_data

    headers_row = table.find("tr")
    if not headers_row:
        return rows_data
    col_names = [th.get_text(strip=True).lower() for th in headers_row.find_all(["th", "td"])]

    for row in table.find_all("tr")[1:]:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue
        cell_texts = [c.get_text(strip=True) for c in cells]
        # Try to map by column name
        row_dict = dict(zip(col_names, cell_texts))

        # Find object / oggetto
        oggetto = ""
        for key in ("oggetto", "titolo", "descrizione", "object"):
            if key in row_dict:
                oggetto = row_dict[key]
                break
        if not oggetto and cell_texts:
            oggetto = " | ".join(cell_texts)

        # Find date
        data_pub = ""
        for key in ("data pubblicazione", "data", "data pubbl.", "datapubblicazione"):
            if key in row_dict:
                data_pub = row_dict[key]
                break

        # Find link
        link = ""
        link_tag = row.find("a", href=True)
        if link_tag:
            href = link_tag["href"]
            if href.startswith("http"):
                link = href
            elif href.startswith("/"):
                # Relative URL - reconstruct from portal base
                link = href

        rows_data.append({"data": data_pub, "oggetto": oggetto, "link": link})

    return rows_data


def _asp_get_next_page(soup, base_url: str) -> str | None:
    """Return URL of next page, or None."""
    # Common patterns: button/link with text "Successiva", ">" or aria-label next
    for a in soup.find_all("a", href=True):
        txt = a.get_text(strip=True).lower()
        aria = (a.get("aria-label") or "").lower()
        if txt in (">", ">>", "successiva", "next", "avanti") or "next" in aria:
            href = a["href"]
            if href.startswith("http"):
                return href
            elif href.startswith("/"):
                from urllib.parse import urlparse
                parsed = urlparse(base_url)
                return f"{parsed.scheme}://{parsed.netloc}{href}"
    # Look for pagination with page number POST param
    for btn in soup.find_all("button"):
        txt = btn.get_text(strip=True).lower()
        if txt in (">", "successiva", "next"):
            onclick = btn.get("onclick", "")
            page_match = re.search(r"page[=\(](\d+)", onclick, re.I)
            if page_match:
                return page_match.group(1)
    return None


def scrape_asp(portal_name: str, base_url: str) -> tuple[list[dict], str | None]:
    """
    Attempt to scrape an ASP AlboOnline portal.
    Returns (list_of_matching_rows, error_message_or_None).
    """
    session = make_session()
    results = []

    try:
        # Step 1: GET the search page to grab form tokens
        r0 = session.get(base_url, timeout=20, verify=False)
        if r0.status_code != 200:
            return [], f"Portale non raggiungibile (HTTP {r0.status_code})"

        soup0 = BeautifulSoup(r0.text, "lxml")
        hidden = _asp_get_hidden(soup0)

        # Find the form action
        form = soup0.find("form")
        action = base_url
        if form:
            form_action = form.get("action", "")
            if form_action:
                if form_action.startswith("http"):
                    action = form_action
                else:
                    from urllib.parse import urlparse, urljoin
                    action = urljoin(base_url, form_action)

        # Build POST data
        data = dict(hidden)
        # Set date field - try common field IDs/names
        for field_id in ("dataPubblicazioneDal", "dataDal", "dataInizio",
                         "pubblicazioneDal", "dataFrom"):
            data[field_id] = DATE_FROM_2D

        # Step 2: POST search
        page_num = 0
        max_pages = 50
        current_url = action

        while page_num < max_pages:
            try:
                r = session.post(current_url, data=data, timeout=20, verify=False)
            except Exception as e:
                return results, f"Errore connessione: {e}"

            if r.status_code != 200:
                if not results:
                    return [], f"Portale non raggiungibile (HTTP {r.status_code})"
                break

            soup = BeautifulSoup(r.text, "lxml")
            page_rows = _asp_parse_rows(soup)

            if not page_rows:
                break

            for row in page_rows:
                if matches_keywords(row["oggetto"]):
                    results.append(row)

            next_url = _asp_get_next_page(soup, current_url)
            if not next_url:
                break
            current_url = next_url
            page_num += 1
            time.sleep(0.5)

        return results, None

    except requests.exceptions.Timeout:
        return [], "Portale non raggiungibile (timeout connessione)"
    except requests.exceptions.ConnectionError as e:
        return [], f"Portale non raggiungibile (errore connessione: {e})"
    except Exception as e:
        return [], f"Errore imprevisto: {e}"


# ── Regione Calabria ───────────────────────────────────────────────────────────
RC_BASE = "https://www.regione.calabria.it/provvedimenti-della-regione/"


def _rc_parse_page(soup) -> list[dict]:
    """Parse one page of Regione Calabria results."""
    rows_data = []
    table = soup.find("table", class_=re.compile("table-striped", re.I))
    if not table:
        return rows_data

    for row in table.find_all("tr")[1:]:  # skip header
        cells = row.find_all(["td", "th"])
        if len(cells) < 5:
            continue
        # Col 0=Tipologia, 1=Data, 2=N, 3=Dipartimento, 4=Oggetto, 5=Dettaglio
        data_pub = cells[1].get_text(strip=True)
        oggetto  = cells[4].get_text(strip=True)
        link_tag = cells[5].find("a") if len(cells) > 5 else None
        link     = link_tag.get("href", "") if link_tag else ""
        rows_data.append({"data": data_pub, "oggetto": oggetto, "link": link})

    return rows_data


def _rc_max_page(soup) -> int:
    """Find the last page number from pagination links."""
    max_p = 1
    for a in soup.find_all("a", href=re.compile(r"paged=\d+")):
        m = re.search(r"paged=(\d+)", a.get("href", ""))
        if m:
            max_p = max(max_p, int(m.group(1)))
    return max_p


def scrape_regione_calabria() -> tuple[list[dict], str | None]:
    """Scrape all pages from Regione Calabria provvedimenti."""
    session = make_session()
    results = []

    try:
        # Page 1
        params = {
            "filter_date_from": DATE_FROM_2D,
            "filter_active": "true",
            "sort_order": "date_asc",
            "paged": "1",
        }
        r = session.get(RC_BASE, params=params, timeout=30, verify=False)
        if r.status_code != 200:
            return [], f"Portale non raggiungibile (HTTP {r.status_code})"

        soup = BeautifulSoup(r.text, "lxml")
        rows = _rc_parse_page(soup)
        for row in rows:
            if matches_keywords(row["oggetto"]):
                results.append(row)

        max_page = _rc_max_page(soup)
        print(f"  Regione Calabria: {max_page} pagine trovate, {len(rows)} righe p.1")

        for page in range(2, max_page + 1):
            params["paged"] = str(page)
            try:
                r = session.get(RC_BASE, params=params, timeout=30, verify=False)
            except Exception as e:
                print(f"  Errore pagina {page}: {e}")
                break
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, "lxml")
            rows = _rc_parse_page(soup)
            if not rows:
                break
            for row in rows:
                if matches_keywords(row["oggetto"]):
                    results.append(row)
            time.sleep(0.3)

        return results, None

    except requests.exceptions.Timeout:
        return [], "Portale non raggiungibile (timeout)"
    except Exception as e:
        return [], f"Errore: {e}"


# ── TAR Catanzaro ──────────────────────────────────────────────────────────────
TAR_BASE   = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
TAR_PREFIX = (
    "_it_indra_ga_institutional_area_"
    "JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"
)
TAR_PORTLET = (
    "it_indra_ga_institutional_area_"
    "JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe"
)


def _tar_parse_table(soup) -> list[dict]:
    """Parse TAR result table. Columns: NRG Sezione Parte Tipo-Udienza Data-Ud Num-Prov Data-Pub Tipo-Prov Relatore Presidente Esito."""
    rows_data = []
    table = soup.find("table", class_=re.compile("dataTable-hearings", re.I))
    if not table:
        return rows_data

    for row in table.find_all("tr")[1:]:
        cells = row.find_all(["td", "th"])
        if len(cells) < 7:
            continue
        parte    = cells[2].get_text(strip=True)   # "Parte"
        data_pub = cells[6].get_text(strip=True)   # "Data pubblicazione"
        num_prov = cells[5].get_text(strip=True)   # "Numero Provvedimento"
        tipo_prov = cells[7].get_text(strip=True)  # "Tipo Provvedimento"

        # Build oggetto from available fields
        oggetto = f"{tipo_prov} – Parte: {parte} – N.{num_prov}"

        # Try to find link
        link_tag = row.find("a", href=True)
        link     = link_tag["href"] if link_tag else ""
        if link and not link.startswith("http"):
            link = "https://www.giustizia-amministrativa.it" + link

        rows_data.append({"data": data_pub, "oggetto": oggetto, "parte": parte, "link": link})

    return rows_data


def scrape_tar() -> tuple[list[dict], str | None]:
    """Scrape TAR Catanzaro provvedimenti."""
    session = make_session()
    results = []

    try:
        # GET to get session + formDate token + p_auth
        r0 = session.get(TAR_BASE, timeout=20, verify=False)
        if r0.status_code != 200:
            return [], f"Portale non raggiungibile (HTTP {r0.status_code})"

        soup0 = BeautifulSoup(r0.text, "lxml")

        # Check for portlet error on initial page
        portlet_error = soup0.find(string=re.compile(r"Si è verificato un errore", re.I))

        form_date_inp = soup0.find("input", {"name": TAR_PREFIX + "formDate"})
        form_date_val = form_date_inp["value"] if form_date_inp else str(int(time.time() * 1000))

        p_auth_m = re.search(r"p_auth=([A-Za-z0-9]+)", r0.text)
        p_auth   = p_auth_m.group(1) if p_auth_m else ""

        action_url = (
            f"https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
            f"?p_p_id={TAR_PORTLET}"
            f"&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view"
            f"&_{TAR_PORTLET}_javax.portlet.action=%2Fadministrative-acts%2Fsearch"
            f"&p_auth={p_auth}"
        )

        data = {
            TAR_PREFIX + "formDate":        form_date_val,
            TAR_PREFIX + "year":            str(TODAY.year),
            TAR_PREFIX + "number":          "",
            TAR_PREFIX + "hearingDateFrom": "",
            TAR_PREFIX + "hearingDateTo":   "",
            TAR_PREFIX + "section":         "",
            TAR_PREFIX + "type":            "",
            TAR_PREFIX + "specific":        "",
            TAR_PREFIX + "publishDateFrom": DATE_FROM_3D,
            TAR_PREFIX + "publishDateTo":   "",
            TAR_PREFIX + "president":       "",
            TAR_PREFIX + "draftingJudge":   "",
        }

        r = session.post(action_url, data=data, timeout=30, verify=False)
        soup = BeautifulSoup(r.text, "lxml")

        # Check if portlet is in error state
        if soup.find(string=re.compile(r"Si è verificato un errore", re.I)):
            err_msg = (
                "Portlet TAR Catanzaro in stato di errore (backend non disponibile). "
                "Il servizio risulta momentaneamente non raggiungibile. "
                "Verificare manualmente: https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
            )
            return [], err_msg

        rows = _tar_parse_table(soup)
        print(f"  TAR: {len(rows)} righe trovate nella pagina")

        for row in rows:
            # Filter on "Parte" column + oggetto (composed string)
            if matches_keywords(row.get("parte", "")) or matches_keywords(row["oggetto"]):
                results.append(row)

        # Handle pagination (TAR uses DataTables, check for pagination links)
        pag = soup.find_all("a", class_=re.compile(r"paginate|next", re.I))
        # For simplicity, iterate page numbers via POST
        total_txt = soup.find(string=re.compile(r"(\d+)\s*(di|of|record)", re.I))
        if total_txt:
            m = re.search(r"(\d+)\s*(di|of|record)", total_txt)
            if m:
                total = int(m.group(1))
                per_page = 10
                pages = (total + per_page - 1) // per_page
                print(f"  TAR: totale {total} righe, {pages} pagine")
                for page in range(2, min(pages + 1, 51)):
                    data_page = dict(data)
                    data_page[TAR_PREFIX + "delta"] = str(per_page)
                    data_page[TAR_PREFIX + "cur"]   = str(page)
                    try:
                        rp = session.post(action_url, data=data_page, timeout=30, verify=False)
                        soupp = BeautifulSoup(rp.text, "lxml")
                        page_rows = _tar_parse_table(soupp)
                        if not page_rows:
                            break
                        for row in page_rows:
                            if matches_keywords(row.get("parte", "")) or matches_keywords(row["oggetto"]):
                                results.append(row)
                    except Exception as e:
                        print(f"  TAR pagina {page} errore: {e}")
                        break
                    time.sleep(0.5)

        return results, None

    except requests.exceptions.Timeout:
        return [], "Portale non raggiungibile (timeout)"
    except Exception as e:
        return [], f"Errore: {e}"


# ── Excel output ───────────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT   = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
LINK_FONT     = Font(color="0563C1", underline="single", name="Calibri", size=10)
NORMAL_FONT   = Font(name="Calibri", size=10)
ERROR_FONT    = Font(color="C00000", bold=True, name="Calibri", size=10)
NOTE_FILL     = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL      = PatternFill("solid", fgColor="D6E4F0")

SHEET_COLS = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
COL_WIDTHS  = [15,     60,        55,                             50]


def _add_header_row(ws):
    for col_idx, (title, width) in enumerate(zip(SHEET_COLS, COL_WIDTHS), start=1):
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20


def _add_data_rows(ws, results: list[dict], start_row: int = 2):
    for i, row in enumerate(results):
        r = start_row + i
        fill = ALT_FILL if i % 2 == 0 else PatternFill()

        oggetto = row.get("oggetto", "")
        data    = row.get("data",    "")
        link    = row.get("link",    "")
        desc    = (oggetto[:150] + "…") if len(oggetto) > 150 else oggetto

        ws.cell(r, 1, data).font      = NORMAL_FONT
        ws.cell(r, 1).fill             = fill
        ws.cell(r, 1).alignment        = Alignment(vertical="center")

        ws.cell(r, 2, oggetto).font   = NORMAL_FONT
        ws.cell(r, 2).fill             = fill
        ws.cell(r, 2).alignment        = Alignment(wrap_text=True, vertical="top")

        ws.cell(r, 3, desc).font      = NORMAL_FONT
        ws.cell(r, 3).fill             = fill
        ws.cell(r, 3).alignment        = Alignment(wrap_text=True, vertical="top")

        # Hyperlink in col 4
        link_cell = ws.cell(r, 4, link if link else "N/D")
        link_cell.fill      = fill
        link_cell.alignment = Alignment(vertical="center")
        if link:
            link_cell.hyperlink = link
            link_cell.font      = LINK_FONT
        else:
            link_cell.font = NORMAL_FONT

        ws.row_dimensions[r].height = 45


def _write_error_row(ws, message: str):
    ws.merge_cells("A2:D2")
    cell = ws.cell(2, 1, message)
    cell.font      = ERROR_FONT
    cell.fill      = NOTE_FILL
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 60


def _write_no_results(ws):
    ws.merge_cells("A2:D2")
    cell = ws.cell(2, 1, "Nessun risultato per il periodo selezionato")
    cell.font      = Font(italic=True, color="595959", name="Calibri", size=10)
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 30


def build_excel(sheet_data: dict[str, tuple]) -> BytesIO:
    """
    sheet_data = {
        sheet_name: (results_list, error_or_None)
    }
    Returns BytesIO of the xlsx file.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sheet_name, (results, error) in sheet_data.items():
        ws = wb.create_sheet(title=sheet_name)
        _add_header_row(ws)

        if error:
            _write_error_row(ws, f"⚠ {error}")
        elif not results:
            _write_no_results(ws)
        else:
            _add_data_rows(ws, results)

        # Freeze top row
        ws.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    run_date = TODAY.strftime("%d-%m-%Y")
    filename = f"Monitoraggio_Atti_Calabria_{run_date}.xlsx"

    print(f"\n{'='*60}")
    print(f"  MONITORAGGIO ATTI CALABRIA – {TODAY.strftime('%d/%m/%Y')}")
    print(f"  Periodo: dal {DATE_FROM_2D} ad oggi")
    print(f"{'='*60}\n")

    sheet_data = {}

    # ── 1) ASP portals ─────────────────────────────────────────────────────────
    for portal_name, base_url in ASP_PORTALS.items():
        print(f"[{portal_name}] scraping {base_url} ...")
        results, error = scrape_asp(portal_name, base_url)
        if error:
            print(f"  ✗ {error}")
        else:
            print(f"  ✓ {len(results)} match trovati")
        sheet_data[portal_name] = (results, error)

    # ── 2) Regione Calabria ────────────────────────────────────────────────────
    print(f"\n[Regione Calabria] scraping {RC_BASE} ...")
    rc_results, rc_error = scrape_regione_calabria()
    if rc_error:
        print(f"  ✗ {rc_error}")
    else:
        print(f"  ✓ {len(rc_results)} match trovati")
    sheet_data["Regione Calabria"] = (rc_results, rc_error)

    # ── 3) TAR Catanzaro ───────────────────────────────────────────────────────
    print(f"\n[TAR Catanzaro] scraping …")
    tar_results, tar_error = scrape_tar()
    if tar_error:
        print(f"  ✗ {tar_error}")
    else:
        print(f"  ✓ {len(tar_results)} match trovati")
    sheet_data["TAR Catanzaro"] = (tar_results, tar_error)

    # ── Build Excel ────────────────────────────────────────────────────────────
    print(f"\nGenerazione file Excel: {filename}")
    xlsx_buf = build_excel(sheet_data)

    # Save locally
    local_path = f"/home/user/NS/output/{filename}"
    import os
    os.makedirs("/home/user/NS/output", exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(xlsx_buf.getvalue())
    print(f"Salvato: {local_path}")

    # Summary
    print(f"\n{'='*60}")
    print("  RIEPILOGO RISULTATI")
    print(f"{'='*60}")
    for name, (results, error) in sheet_data.items():
        if error:
            print(f"  {name:30s}  ✗  {error[:70]}")
        else:
            print(f"  {name:30s}  ✓  {len(results):3d} match")
    print(f"{'='*60}\n")

    return local_path, xlsx_buf, filename


if __name__ == "__main__":
    local_path, xlsx_buf, filename = main()
    print(f"Done. File: {local_path}")
