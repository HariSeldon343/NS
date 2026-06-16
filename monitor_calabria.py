#!/usr/bin/env python3
"""
Monitor atti amministrativi Calabria
Portali: ASP (CO, CZ, KR, RC, VV), Regione Calabria, TAR Catanzaro
Output: XLSX con un foglio per portale
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
import re
import time
import urllib3
import urllib.parse

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─── Date range ──────────────────────────────────────────────────────────────
TODAY       = datetime.now()
DATE_FROM   = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM3  = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")
DATE_TODAY  = TODAY.strftime("%Y-%m-%d")
DATE_IT     = TODAY.strftime("%d-%m-%Y")  # per filename
DATE_FROM_IT = (TODAY - timedelta(days=2)).strftime("%d/%m/%Y")

print(f"[*] Esecuzione: {TODAY.strftime('%d/%m/%Y %H:%M')}")
print(f"[*] Intervallo 2gg: {DATE_FROM} → {DATE_TODAY}")
print(f"[*] Intervallo 3gg (TAR): {DATE_FROM3} → {DATE_TODAY}")

# ─── Keywords ────────────────────────────────────────────────────────────────
KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia", "Life",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]

def matches_keywords(text: str) -> bool:
    if not text:
        return False
    for kw in KEYWORDS:
        if kw.upper() == "LIFE":
            if re.search(r'\bLIFE\b', text, re.IGNORECASE):
                return True
        else:
            if kw.upper() in text.upper():
                return True
    return False

# ─── HTTP helpers ─────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8",
    "Connection": "keep-alive",
}

def make_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# ═══════════════════════════════════════════════════════════════════════════
# REGIONE CALABRIA
# ═══════════════════════════════════════════════════════════════════════════

def fetch_regione_calabria() -> list:
    """
    Portale: https://www.regione.calabria.it/provvedimenti-della-regione/
    Metodo:  GET con ?filter_date_from=YYYY-MM-DD&filter_date_to=YYYY-MM-DD&paged=N
    Tabella: Tipologia | Data Repertoriazione | (N) | Dipartimento | Oggetto | Dettaglio(link)
    """
    print("\n[Regione Calabria] Avvio raccolta atti...")
    results = []
    session = make_session()
    base_url = "https://www.regione.calabria.it/provvedimenti-della-regione/"

    page = 1
    while True:
        params = {
            "filter_date_from": DATE_FROM,
            "filter_date_to":   DATE_TODAY,
            "paged":            str(page),
        }
        try:
            resp = session.get(base_url, params=params, timeout=30, verify=False)
            resp.raise_for_status()
        except Exception as e:
            print(f"  [!] Errore pagina {page}: {e}")
            break

        soup = BeautifulSoup(resp.text, "lxml")
        table = soup.find("table")
        if not table:
            print(f"  Pagina {page}: nessuna tabella trovata")
            break

        rows = table.find_all("tr")
        data_rows = rows[1:]  # skip header
        if not data_rows:
            print(f"  Pagina {page}: tabella vuota")
            break

        matched = 0
        for row in data_rows:
            cells = row.find_all(["td", "th"])
            if len(cells) < 3:
                continue

            date_text = cells[1].get_text(strip=True) if len(cells) > 1 else ""

            # Oggetto: second-to-last substantive cell (before the → link cell)
            # Row structure (5 cells): Tipologia | Data | Dipartimento | Oggetto | →
            oggetto = ""
            if len(cells) >= 5:
                oggetto = cells[3].get_text(strip=True)
            elif len(cells) == 4:
                oggetto = cells[2].get_text(strip=True)
            elif len(cells) == 3:
                oggetto = cells[2].get_text(strip=True)

            # Link: find first <a href> that looks like a detail URL
            link = ""
            for c in cells:
                a = c.find("a", href=True)
                if a and ("provvediment" in a["href"] or "decreto" in a["href"] or "deliber" in a["href"]):
                    link = a["href"]
                    break
            if not link:
                for c in cells:
                    a = c.find("a", href=True)
                    if a and a["href"].startswith("http"):
                        link = a["href"]
                        break

            if matches_keywords(oggetto):
                results.append({
                    "data":    date_text,
                    "oggetto": oggetto,
                    "link":    link,
                })
                matched += 1

        print(f"  Pagina {page}: {len(data_rows)} righe, {matched} match")

        # Pagination: check if next page link exists
        has_next = any(
            f"paged={page+1}" in (a.get("href") or "")
            for a in soup.find_all("a", href=True)
        )
        if not has_next:
            break
        page += 1
        time.sleep(0.4)

    print(f"  TOTALE Regione Calabria: {len(results)} atti trovati")
    return results


# ═══════════════════════════════════════════════════════════════════════════
# TAR CATANZARO
# ═══════════════════════════════════════════════════════════════════════════

TAR_PORTLET_NS = (
    "_it_indra_ga_institutional_area_"
    "JurisdictionalActivityAdministrativeActsWebPortlet_"
    "INSTANCE_jjYpzZYF4Qfe_"
)
TAR_BASE = "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"

def fetch_tar_catanzaro() -> list:
    """
    Portale: giustizia-amministrativa.it
    Metodo:  POST form Liferay (portlet lifecycle=1)
    Filtro:  publishDateFrom/To (ultimi 3 giorni)
    Colonne risultati: NRG | Sezione | Parte | Tipo Udienza | Data Udienza |
                       Numero Provvedimento | Data pubblicazione | Tipo Provvedimento |
                       Relatore | Presidente | Esito
    Match su 'Parte'
    """
    print("\n[TAR Catanzaro] Avvio raccolta atti...")
    results = []
    session = make_session()

    # Step 1: GET per ottenere session cookie + p_auth + formDate
    try:
        r0 = session.get(TAR_BASE, timeout=30, verify=False)
        r0.raise_for_status()
    except Exception as e:
        print(f"  [!] Impossibile raggiungere TAR: {e}")
        return [{"data": "", "oggetto": f"ERRORE CONNESSIONE: {e}", "link": "", "_error": True}]

    soup0 = BeautifulSoup(r0.text, "lxml")

    # Find form
    form = None
    for f in soup0.find_all("form"):
        if "administrative-acts" in (f.get("action") or ""):
            form = f
            break

    if not form:
        print("  [!] Form non trovato nel portale TAR")
        return [{"data": "", "oggetto": "ERRORE: form di ricerca non trovato", "link": "", "_error": True}]

    action = form["action"]

    # Build base form data (all inputs, including those in disabled fieldset)
    form_data = {}
    for inp in form.find_all("input"):
        n = inp.get("name")
        v = inp.get("value", "")
        if n:
            form_data[n] = v
    for sel in form.find_all("select"):
        n = sel.get("name")
        if n:
            form_data[n] = ""

    # Set date range (3 giorni per TAR)
    form_data[TAR_PORTLET_NS + "publishDateFrom"] = DATE_FROM3
    form_data[TAR_PORTLET_NS + "publishDateTo"]   = DATE_TODAY

    session.headers["Referer"] = TAR_BASE

    # Step 2: POST search
    page_delta = 0
    while True:
        post_data = dict(form_data)
        if page_delta > 0:
            post_data[TAR_PORTLET_NS + "delta"] = "50"
            post_data[TAR_PORTLET_NS + "cur"]   = str(page_delta)

        try:
            r1 = session.post(action, data=post_data, timeout=30, verify=False)
            r1.raise_for_status()
        except Exception as e:
            print(f"  [!] POST TAR fallito: {e}")
            break

        soup1 = BeautifulSoup(r1.text, "lxml")

        # Detect backend error
        err_text = soup1.find(string=re.compile(r'Si è verificato un errore', re.I))
        if err_text:
            print(f"  [!] Errore backend TAR: '{str(err_text).strip()[:80]}'")
            print("      Nota: il portale TAR restituisce errore dal server per ricerche remote.")
            return [{"data": "", "oggetto": "ERRORE SERVER TAR: il portale non risponde alle ricerche da rete esterna. Consultare manualmente: " + TAR_BASE, "link": TAR_BASE, "_error": True}]

        # Parse results table
        table = soup1.find("table")
        if not table:
            print(f"  Nessuna tabella TAR trovata (pagina {page_delta})")
            break

        tar_rows = table.find_all("tr")
        if len(tar_rows) <= 1:
            print(f"  TAR pagina {page_delta}: tabella vuota")
            break

        # Headers: NRG | Sezione | Parte | Tipo Udienza | Data Udienza | N.Prov | Data pubbl | Tipo Prov | Relatore | Presidente | Esito
        headers = [th.get_text(strip=True).lower() for th in tar_rows[0].find_all(["th", "td"])]

        def col_i(names):
            for n in names:
                for i, h in enumerate(headers):
                    if n in h:
                        return i
            return None

        idx_parte  = col_i(["parte"])
        idx_data   = col_i(["data pubbl", "pubblicazione"])
        idx_nrg    = col_i(["nrg"])
        idx_tipo   = col_i(["tipo prov", "specifica"])
        idx_num    = col_i(["numero prov", "n.prov", "num"])

        matched = 0
        for row in tar_rows[1:]:
            cells = row.find_all(["td", "th"])
            if not cells:
                continue

            def cell_text(idx):
                if idx is not None and idx < len(cells):
                    return cells[idx].get_text(strip=True)
                return ""

            parte     = cell_text(idx_parte)
            data_pub  = cell_text(idx_data)
            nrg       = cell_text(idx_nrg)
            tipo_prov = cell_text(idx_tipo)
            num_prov  = cell_text(idx_num)

            # Componi oggetto descrittivo
            oggetto_parts = filter(None, [nrg, parte, tipo_prov, num_prov])
            oggetto = " | ".join(oggetto_parts)

            link = ""
            for c in cells:
                a = c.find("a", href=True)
                if a:
                    href = a["href"]
                    if not href.startswith("http"):
                        href = "https://www.giustizia-amministrativa.it" + href
                    link = href
                    break

            if matches_keywords(parte) or matches_keywords(oggetto):
                results.append({
                    "data":    data_pub,
                    "oggetto": oggetto,
                    "link":    link,
                })
                matched += 1

        print(f"  TAR pagina {page_delta}: {len(tar_rows)-1} righe, {matched} match")

        # Next page
        next_link = soup1.find("a", class_=re.compile(r'next|successiv'))
        if not next_link:
            # Look for page number links beyond current
            pager = soup1.find(class_=re.compile(r'pagination|paginator|nav-links'))
            if pager:
                cur_links = pager.find_all("a", href=True)
                advanced = [a for a in cur_links if "cur=" in a.get("href", "")]
                if advanced:
                    page_delta += 1
                    time.sleep(0.5)
                    continue
            break
        page_delta += 1
        time.sleep(0.5)

    print(f"  TOTALE TAR: {len(results)} atti trovati")
    return results


# ═══════════════════════════════════════════════════════════════════════════
# ASP PORTALS (SISR)
# ═══════════════════════════════════════════════════════════════════════════

ASP_PORTALS = {
    "ASP Cosenza":         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def fetch_asp_portal(name: str, base_url: str) -> list:
    """
    Portali SISR AlboOnline. Richiedono accesso da rete intranet Regione Calabria.
    Da reti esterne (cloud, VPS) restituiscono timeout/503.
    """
    print(f"\n[{name}] Connessione a {base_url}")
    results = []
    session = make_session()

    # Step 1: GET iniziale per session + hidden fields + CSRF
    try:
        r0 = session.get(base_url, timeout=20, verify=False)
        r0.raise_for_status()
    except requests.exceptions.Timeout:
        print(f"  [!] Timeout: portale non raggiungibile da rete esterna")
        return [{"data": "", "oggetto": f"NON ACCESSIBILE: timeout connessione. Il portale {name} è raggiungibile solo dalla rete intranet della Regione Calabria. URL: {base_url}", "link": base_url, "_error": True}]
    except requests.exceptions.ConnectionError as e:
        print(f"  [!] Errore connessione: {e}")
        return [{"data": "", "oggetto": f"NON ACCESSIBILE: errore connessione ({type(e).__name__}). URL: {base_url}", "link": base_url, "_error": True}]
    except Exception as e:
        status = getattr(r0, 'status_code', 'N/A') if 'r0' in dir() else 'N/A'
        print(f"  [!] Errore HTTP {status}: {e}")
        return [{"data": "", "oggetto": f"NON ACCESSIBILE: errore HTTP {status}. URL: {base_url}", "link": base_url, "_error": True}]

    soup0 = BeautifulSoup(r0.text, "lxml")
    parsed = urllib.parse.urlparse(base_url)
    site_base = f"{parsed.scheme}://{parsed.netloc}"

    # Collect form hidden fields and action
    hidden = {}
    form = soup0.find("form")
    if form:
        for inp in form.find_all("input", {"type": "hidden"}):
            n = inp.get("name") or inp.get("id")
            v = inp.get("value", "")
            if n:
                hidden[n] = v
        action_raw = form.get("action", base_url)
        if action_raw and not action_raw.startswith("http"):
            action = site_base + action_raw if action_raw.startswith("/") else base_url
        else:
            action = action_raw or base_url
    else:
        action = base_url

    print(f"  Form action: {action}")

    # Iterate pages
    page = 0
    total_scraped = 0
    while True:
        post_data = dict(hidden)
        post_data.update({
            "dataPubblicazioneDal": DATE_FROM,
            "dataPubblicazioneAl":  DATE_TODAY,
            "page":                 str(page),
            "rows":                 "50",
            "pageSize":             "50",
            "pageIndex":            str(page),
        })

        try:
            r1 = session.post(action, data=post_data, timeout=25, verify=False)
            r1.raise_for_status()
        except requests.exceptions.Timeout:
            if page == 0:
                return [{"data": "", "oggetto": f"NON ACCESSIBILE: timeout POST. URL: {base_url}", "link": base_url, "_error": True}]
            break
        except Exception as e:
            if page == 0:
                return [{"data": "", "oggetto": f"NON ACCESSIBILE: errore POST ({e}). URL: {base_url}", "link": base_url, "_error": True}]
            break

        soup1 = BeautifulSoup(r1.text, "lxml")
        rows_found, matched = parse_asp_table(soup1, site_base, base_url, results)
        total_scraped += rows_found
        print(f"  Pagina {page}: {rows_found} righe, {matched} match")

        if rows_found == 0:
            if page == 0:
                # Fallback: try GET
                return fetch_asp_get_fallback(session, base_url, name, site_base, hidden)
            break

        # Next page
        if not find_asp_next_page(soup1, page):
            break
        page += 1
        time.sleep(0.4)

    print(f"  TOTALE {name}: {len(results)} atti trovati")
    return results


def fetch_asp_get_fallback(session, base_url, name, site_base, hidden) -> list:
    """Fallback GET-based pagination for ASP portals."""
    results = []
    page = 0
    while True:
        params = {
            "dataPubblicazioneDal": DATE_FROM,
            "dataPubblicazioneAl":  DATE_TODAY,
            "page":  str(page),
            "rows":  "50",
        }
        try:
            r = session.get(base_url, params=params, timeout=20, verify=False)
            r.raise_for_status()
        except Exception as e:
            if page == 0:
                return [{"data": "", "oggetto": f"NON ACCESSIBILE (GET fallback): {e}. URL: {base_url}", "link": base_url, "_error": True}]
            break
        soup = BeautifulSoup(r.text, "lxml")
        rows_found, matched = parse_asp_table(soup, site_base, base_url, results)
        print(f"  GET p{page}: {rows_found} righe, {matched} match")
        if rows_found == 0:
            break
        if not find_asp_next_page(soup, page):
            break
        page += 1
        time.sleep(0.4)
    return results


def parse_asp_table(soup: BeautifulSoup, site_base: str, base_url: str, results: list):
    """Parse result table from ASP AlboOnline portal. Returns (total_rows, matched)."""
    # Try known table IDs/classes first, then any table with multiple rows
    table = None
    for candidate_id in ["tableAlbo", "risultati", "tabella", "alboTable", "listaAlbo"]:
        t = soup.find("table", {"id": candidate_id})
        if not t:
            t = soup.find("table", class_=candidate_id)
        if t:
            table = t
            break
    if not table:
        for t in soup.find_all("table"):
            if len(t.find_all("tr")) > 1:
                table = t
                break
    if not table:
        return 0, 0

    rows = table.find_all("tr")
    if len(rows) < 2:
        return 0, 0

    # Detect headers
    hdr_cells = rows[0].find_all(["th", "td"])
    headers = [c.get_text(strip=True).lower() for c in hdr_cells]

    def col_i(names):
        for n in names:
            for i, h in enumerate(headers):
                if n in h:
                    return i
        return None

    idx_date    = col_i(["data pubbl", "data di pubbl", "data pubblica", "pubblicazione", "data"])
    idx_oggetto = col_i(["oggetto", "titolo", "descrizione"])
    idx_num     = col_i(["numero", "num.", "n."])

    matched = 0
    for row in rows[1:]:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue

        def cell_text(idx):
            if idx is not None and idx < len(cells):
                return cells[idx].get_text(strip=True)
            return ""

        date_text = cell_text(idx_date)
        if not date_text:
            # search all cells for date pattern
            for c in cells:
                if re.search(r'\d{2}[/.\-]\d{2}[/.\-]\d{4}', c.get_text()):
                    date_text = c.get_text(strip=True)
                    break

        oggetto = cell_text(idx_oggetto)
        if not oggetto:
            # longest text cell
            max_l = 0
            for c in cells:
                t = c.get_text(strip=True)
                if len(t) > max_l and not re.match(r'^\d{2}[/.\-]\d{2}[/.\-]\d{4}$', t):
                    max_l = len(t)
                    oggetto = t

        # Find link
        link = ""
        for c in cells:
            a = c.find("a", href=True)
            if a:
                href = a["href"]
                if not href.startswith("http"):
                    href = site_base + href if href.startswith("/") else base_url + "/" + href
                link = href
                break

        if matches_keywords(oggetto):
            results.append({"data": date_text, "oggetto": oggetto, "link": link})
            matched += 1

    return len(rows) - 1, matched


def find_asp_next_page(soup: BeautifulSoup, current_page: int) -> bool:
    """Return True if there's a next page in ASP portal."""
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True).lower()
        if text in ["successiva", "next", ">", "»", "avanti"]:
            return True
    pager = soup.find(class_=re.compile(r'pager|pagination|pagina'))
    if pager:
        for a in pager.find_all("a", href=True):
            if f"page={current_page+1}" in a["href"] or f"pageIndex={current_page+1}" in a["href"]:
                return True
    return False


# ═══════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════════════

SHEET_ORDER = [
    "ASP Cosenza",
    "ASP Catanzaro",
    "ASP Crotone",
    "ASP Reggio Calabria",
    "ASP Vibo Valentia",
    "Regione Calabria",
    "TAR Catanzaro",
]

def write_excel(all_results: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT   = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    LINK_FONT  = Font(color="0563C1", underline="single", name="Calibri", size=10)
    NORM_FONT  = Font(name="Calibri", size=10)
    ERR_FONT   = Font(name="Calibri", size=10, color="C00000", italic=True)
    WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
    TOP_ALIGN  = Alignment(vertical="top")
    thin_side  = Side(style="thin", color="BFBFBF")
    thin_bord  = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for sheet_name in SHEET_ORDER:
        ws = wb.create_sheet(title=sheet_name[:31])
        rows = all_results.get(sheet_name, [])

        # Header row
        ws.append(["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"])
        ws.row_dimensions[1].height = 22
        for cell in ws[1]:
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = TOP_ALIGN
            cell.border = thin_bord

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 65
        ws.column_dimensions["C"].width = 42
        ws.column_dimensions["D"].width = 16
        ws.freeze_panes = "A2"

        if not rows:
            ws.append(["Nessun risultato", "", "", ""])
            ws["A2"].font = Font(italic=True, color="808080", name="Calibri", size=10)
            continue

        # Check if all rows are errors
        all_errors = all(r.get("_error") for r in rows)

        for r in rows:
            data       = r.get("data", "")
            oggetto    = r.get("oggetto", "")
            descrizione = oggetto[:150]
            link       = r.get("link", "")
            is_error   = r.get("_error", False)

            row_num = ws.max_row + 1
            ws.append([data, oggetto, descrizione, ""])

            for col in range(1, 5):
                cell = ws.cell(row=row_num, column=col)
                cell.alignment = WRAP_ALIGN
                cell.border = thin_bord
                if is_error:
                    cell.font = ERR_FONT
                else:
                    cell.font = NORM_FONT

            link_cell = ws.cell(row=row_num, column=4)
            if link and not is_error:
                link_cell.value = "Apri atto"
                link_cell.hyperlink = link
                link_cell.font = LINK_FONT
                link_cell.alignment = TOP_ALIGN
            elif link and is_error:
                link_cell.value = "URL portale"
                link_cell.hyperlink = link
                link_cell.font = Font(color="C00000", underline="single", name="Calibri", size=10, italic=True)
                link_cell.alignment = TOP_ALIGN
            else:
                link_cell.value = "—"

    wb.save(output_path)
    print(f"\n[OK] File salvato: {output_path}")


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    all_results = {}

    # ASP portals
    for portal_name, portal_url in ASP_PORTALS.items():
        all_results[portal_name] = fetch_asp_portal(portal_name, portal_url)

    # Regione Calabria
    all_results["Regione Calabria"] = fetch_regione_calabria()

    # TAR Catanzaro
    all_results["TAR Catanzaro"] = fetch_tar_catanzaro()

    # Summary
    print("\n" + "="*55)
    print(f"{'RIEPILOGO':^55}")
    print("="*55)
    total_match = 0
    total_err   = 0
    for k in SHEET_ORDER:
        rows = all_results.get(k, [])
        err_count = sum(1 for r in rows if r.get("_error"))
        ok_count  = len(rows) - err_count
        if err_count:
            print(f"  {k:<25}  ⚠  non accessibile")
            total_err += 1
        else:
            print(f"  {k:<25}  {ok_count:>3} match")
            total_match += ok_count
    print(f"\n  Totale atti trovati: {total_match}")
    print(f"  Portali non accessibili: {total_err}")
    print("="*55)

    # Output filename
    output_filename = f"Monitoraggio_Atti_Calabria_{DATE_IT}.xlsx"
    output_path = f"/home/user/NS/{output_filename}"
    write_excel(all_results, output_path)
    return output_path


if __name__ == "__main__":
    output = main()
    print(f"\nFile: {output}")
