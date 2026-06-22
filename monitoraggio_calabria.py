#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria - scraper multi-portale
Portali: ASP (CO, CZ, KR, RC, VV), Regione Calabria, TAR Catanzaro
"""

import requests
import json
import re
import warnings
from datetime import date, datetime, timedelta
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ─── Configuration ───────────────────────────────────────────────────────────

TODAY = date(2026, 6, 22)
DATE_FROM_2DAYS = TODAY - timedelta(days=2)   # 2026-06-20
DATE_FROM_3DAYS = TODAY - timedelta(days=3)   # 2026-06-19 (for TAR: 3 days)
CUTOFF_7DAYS = TODAY - timedelta(days=7)       # for Regione Calabria pagination stop

KEYWORDS = [
    "ADI",
    "Assistenza domiciliare",
    "ANMIC",
    "Accreditamento",
    "Aumento di budget",
    "Autismo",
    r"Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione",
    "Autorizzazioni",
    "Budget",
    "Casa Giardino",
    "Centro San Giuseppe",
    "Centro salute e benessere",
    "Fabbisogni LEA",
    "Fisiolab",
    "Fisioterapia",
    r"\bLIFE\b",          # word-boundary for Life
    "Parere commissione",
    r"Presa d'atto verifica",
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

# Precompile patterns (case-insensitive)
KW_PATTERNS = []
for kw in KEYWORDS:
    if kw.startswith(r"\b"):
        KW_PATTERNS.append(re.compile(kw, re.IGNORECASE))
    else:
        KW_PATTERNS.append(re.compile(re.escape(kw), re.IGNORECASE))

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

ASP_PORTALS = {
    'ASP Cosenza':        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria':'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}

TAR_PORTLET_NS = '_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_'
TAR_AJAX_URL = (
    'https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro'
    '?p_p_id=it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe'
    '&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view'
    '&p_p_resource_id=%2Fadministrative-acts%2Fsearch%2Fresults'
    '&p_p_cacheability=cacheLevelPage'
)
TAR_DOC_BASE = 'https://mdp.giustizia-amministrativa.it/visualizza/?nodeRef=&schema=tar_cz&nrg={nrg}&nomeFile={nomeFile}&subDir=Provvedimenti'


# ─── Keyword matching ─────────────────────────────────────────────────────────

def matches_keywords(text: str) -> bool:
    for pat in KW_PATTERNS:
        if pat.search(text):
            return True
    return False


# ─── ASP Portals (AlboOnline) ────────────────────────────────────────────────

def scrape_asp_portal(name: str, base_url: str) -> list[dict]:
    """
    Scrape an ASP AlboOnline portal for acts published in the last 2 days.
    Returns list of {data, oggetto, link} dicts.
    """
    print(f"\n[{name}] Connecting to {base_url}...")
    s = requests.Session()
    s.headers.update(HEADERS)

    try:
        # Initial GET to get the form and cookies
        r0 = s.get(base_url, timeout=30, verify=False)
        if r0.status_code != 200:
            print(f"  [{name}] HTTP {r0.status_code} - portal not available")
            return [{'_error': f'Portale non raggiungibile (HTTP {r0.status_code})'}]

        soup0 = BeautifulSoup(r0.text, 'html.parser')

        # Find form fields
        form = soup0.find('form')
        if not form:
            print(f"  [{name}] Form not found")
            return [{'_error': 'Form non trovato nella pagina'}]

        form_action = form.get('action', base_url)
        if not form_action.startswith('http'):
            from urllib.parse import urljoin
            form_action = urljoin(base_url, form_action)

        # Collect all hidden fields
        post_data = {}
        for inp in form.find_all('input', type='hidden'):
            if inp.get('name'):
                post_data[inp['name']] = inp.get('value', '')

        # Set date filter
        post_data['dataPubblicazioneDal'] = DATE_FROM_2DAYS.strftime('%Y-%m-%d')
        post_data['dataPubblicazioneAl'] = TODAY.strftime('%Y-%m-%d')

        results = []
        page = 1

        while True:
            print(f"  [{name}] Page {page}...")
            post_data['page'] = str(page)

            r = s.post(form_action, data=post_data, timeout=30, verify=False)
            if r.status_code != 200:
                print(f"  [{name}] Page {page} HTTP {r.status_code}")
                break

            soup = BeautifulSoup(r.text, 'html.parser')

            # Look for results table
            table = soup.find('table', id=re.compile(r'albo|result|atti', re.I))
            if not table:
                table = soup.find('table', class_=re.compile(r'albo|result|atti|table', re.I))
            if not table:
                tables = soup.find_all('table')
                table = tables[0] if tables else None

            if not table:
                # Check for "no results" message
                no_res = soup.find(string=re.compile(r'nessun.*(risultato|atto|record)', re.I))
                if no_res or page > 1:
                    break
                print(f"  [{name}] No table found on page {page}")
                break

            rows = table.find_all('tr')[1:]  # skip header
            if not rows:
                break

            found_any = False
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) < 2:
                    continue

                # Extract data
                texts = [c.get_text(separator=' ', strip=True) for c in cells]
                link_tag = row.find('a', href=True)
                link = link_tag['href'] if link_tag else ''
                if link and not link.startswith('http'):
                    from urllib.parse import urljoin
                    link = urljoin(base_url, link)

                # Try to find date and oggetto in cells
                data_str = texts[0] if texts else ''
                oggetto = texts[1] if len(texts) > 1 else ''

                # Look for Oggetto-like column
                for i, t in enumerate(texts):
                    if len(t) > 30:
                        oggetto = t
                        break

                if matches_keywords(oggetto):
                    results.append({
                        'data': data_str,
                        'oggetto': oggetto,
                        'link': link
                    })
                found_any = True

            if not found_any:
                break

            # Check for next page
            next_pg = soup.find('a', string=re.compile(r'successiv|next|>>', re.I))
            if not next_pg:
                break
            page += 1
            if page > 50:
                break

        print(f"  [{name}] Done. {len(results)} keyword matches.")
        return results

    except Exception as e:
        print(f"  [{name}] Error: {e}")
        return [{'_error': f'Errore connessione: {str(e)[:100]}'}]


# ─── Regione Calabria ─────────────────────────────────────────────────────────

def scrape_regione_calabria() -> list[dict]:
    """
    Scrape Regione Calabria provvedimenti.
    Uses GET with paged=N. Filter is broken server-side; we filter client-side.
    """
    print("\n[Regione Calabria] Starting scrape...")
    s = requests.Session()
    s.headers.update(HEADERS)

    results = []
    cutoff_date = DATE_FROM_2DAYS  # items from last 2 days
    pages_scraped = 0

    for page in range(1, 200):
        url = f'https://www.regione.calabria.it/provvedimenti-della-regione/?paged={page}'
        try:
            r = s.get(url, timeout=30, verify=False)
            if r.status_code != 200:
                print(f"  [Regione] Page {page} HTTP {r.status_code}")
                break
        except Exception as e:
            print(f"  [Regione] Page {page} error: {e}")
            break

        soup = BeautifulSoup(r.text, 'html.parser')
        table = soup.find('table', class_='table-striped')
        if not table:
            print(f"  [Regione] No table on page {page}")
            break

        rows = table.find_all('tr')[1:]
        if not rows:
            break

        oldest_on_page = None
        found_in_range = False

        for row in rows:
            cells = row.find_all(['td', 'th'])
            if len(cells) < 5:
                continue

            tipo = cells[0].get_text(strip=True)
            data_str = cells[1].get_text(strip=True)   # DD/MM/YYYY
            num = cells[2].get_text(strip=True)
            dip = cells[3].get_text(strip=True)
            oggetto = cells[4].get_text(separator=' ', strip=True)
            link_tag = cells[5].find('a', href=True) if len(cells) > 5 else None
            link = link_tag['href'] if link_tag else ''

            # Parse date
            try:
                item_date = datetime.strptime(data_str, '%d/%m/%Y').date()
            except:
                item_date = None

            if oldest_on_page is None and item_date:
                oldest_on_page = item_date

            if item_date:
                oldest_on_page = item_date  # track the oldest on this page

            # Date range filter - accept items from last 5 days (to account for publication lag)
            date_ok = True
            if item_date and item_date < (TODAY - timedelta(days=5)):
                date_ok = False
            if item_date and item_date >= (TODAY - timedelta(days=2)):
                found_in_range = True

            if date_ok and matches_keywords(oggetto):
                results.append({
                    'data': data_str,
                    'oggetto': oggetto,
                    'link': link
                })

        pages_scraped += 1
        print(f"  [Regione] Page {page}: {len(rows)} rows, oldest date: {oldest_on_page}")

        # Stop if all items on page are older than 7 days
        if oldest_on_page and oldest_on_page < CUTOFF_7DAYS:
            print(f"  [Regione] Stopping - items older than 7 days")
            break

    print(f"  [Regione] Done. {pages_scraped} pages, {len(results)} keyword matches.")
    return results


# ─── TAR Catanzaro ────────────────────────────────────────────────────────────

def scrape_tar_catanzaro() -> list[dict]:
    """
    Scrape TAR Catanzaro using the DataTables AJAX endpoint.
    Uses publishDateFrom/To for server-side date filtering.
    """
    print("\n[TAR Catanzaro] Starting scrape...")
    NS = TAR_PORTLET_NS
    s = requests.Session()
    s.headers.update(HEADERS)
    s.headers.update({
        'Origin': 'https://www.giustizia-amministrativa.it',
        'Referer': 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro',
    })

    try:
        # Step 1: Get page to get cookies and form tokens
        r0 = s.get('https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro',
                   timeout=30, verify=False)
        if r0.status_code != 200:
            return [{'_error': f'TAR home HTTP {r0.status_code}'}]

        soup0 = BeautifulSoup(r0.text, 'html.parser')
        form = soup0.find('form', action=lambda x: x and 'administrative-acts' in str(x))
        if not form:
            return [{'_error': 'TAR form not found'}]

        form_action = form.get('action')
        form_date_input = form.find('input', {'name': f'{NS}formDate'})
        form_date = form_date_input['value'] if form_date_input else str(int(datetime.now().timestamp() * 1000))

        # Step 2: POST to initialize search (get DataTable init page)
        post_data = {
            f'{NS}formDate': form_date,
            f'{NS}year': '',
            f'{NS}number': '',
            f'{NS}hearingDateFrom': '',
            f'{NS}hearingDateTo': '',
            f'{NS}section': '',
            f'{NS}type': '',
            f'{NS}specific': '',
            f'{NS}publishDateFrom': DATE_FROM_3DAYS.strftime('%Y-%m-%d'),
            f'{NS}publishDateTo': TODAY.strftime('%Y-%m-%d'),
            f'{NS}president': '',
            f'{NS}draftingJudge': '',
        }
        r1 = s.post(form_action, data=post_data, timeout=30, verify=False)
        if r1.status_code != 200:
            return [{'_error': f'TAR search POST HTTP {r1.status_code}'}]

        # Extract additionalInfo from the JavaScript in the response
        ai_match = re.search(r"d\.additionalInfo\s*=\s*'([^']+)'", r1.text)
        additional_info_str = ai_match.group(1) if ai_match else None

        if not additional_info_str:
            # Build it manually
            additional_info_str = json.dumps({
                "schema": "TAR_CATANZARO",
                "type": None,
                "year": "",
                "number": "",
                "hearingDateFrom": None,
                "hearingDateTo": None,
                "publishDateFrom": DATE_FROM_3DAYS.strftime('%Y-%m-%d'),
                "publishDateTo": TODAY.strftime('%Y-%m-%d'),
                "hearingType": None,
                "nrg": None,
                "section": "",
                "provisionSpecification": "",
                "president": "",
                "draftingJudge": "",
                "subjectMatter": None,
                "page": None,
                "size": None,
                "orderBy": None,
                "orderStrategy": None,
                "queryString": None
            })

        print(f"  [TAR] additionalInfo: {additional_info_str[:100]}...")

        # Step 3: AJAX DataTables request - fetch all records
        results = []
        start = 0
        page_size = 100

        while True:
            dt_payload = {
                "draw": (start // page_size) + 1,
                "columns": [
                    {"data": "nrgFascicolo", "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "sezione",       "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "parte",         "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "tipoUdienza",   "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "dataUdienza",   "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "numProvvedimento","name":"", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "dataPubblicazione","name":"","searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "tipoProvvedimento","name":"","searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "relatore",      "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "presidente",    "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                    {"data": "esito",         "name": "", "searchable": True, "orderable": True, "search": {"value": "", "regex": False}},
                ],
                "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
                "start": start,
                "length": page_size,
                "search": {"value": "", "regex": False},
                "additionalInfo": additional_info_str,
            }

            r_ajax = s.post(
                TAR_AJAX_URL,
                data=json.dumps(dt_payload),
                headers={
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                },
                timeout=30,
                verify=False
            )

            if r_ajax.status_code != 200:
                print(f"  [TAR] AJAX HTTP {r_ajax.status_code}")
                break

            try:
                data = r_ajax.json()
            except Exception as e:
                print(f"  [TAR] JSON parse error: {e}")
                break

            records = data.get('data', [])
            total = data.get('recordsFiltered', 0)
            print(f"  [TAR] Page {(start//page_size)+1}: {len(records)} records (total={total})")

            for rec in records:
                parte = rec.get('parte', '') or ''
                tipo_prov = rec.get('tipoProvvedimento', '') or ''
                data_pub = rec.get('dataPubblicazione', '') or ''
                num_prov = rec.get('numProvvedimento', '') or ''
                nrg = rec.get('nrgFascicolo', '') or ''
                nome_file = rec.get('nomeFile', '') or ''
                esito = rec.get('esito', '') or ''

                # The "Oggetto" for TAR is the Parte field
                # (many are anonymized as XXX_OMISSIS_XXX)
                # Build a descriptive text combining available info
                oggetto = f"{parte} - {tipo_prov} N.{num_prov} ({data_pub})"
                if esito.strip():
                    oggetto += f" - Esito: {esito.strip()}"

                # Build link
                link = ''
                if nome_file and nrg:
                    link = TAR_DOC_BASE.format(nrg=nrg, nomeFile=nome_file)

                # Match keywords against parte field
                if matches_keywords(parte):
                    results.append({
                        'data': data_pub,
                        'oggetto': oggetto,
                        'parte': parte,
                        'link': link
                    })

            if start + page_size >= total:
                break
            start += page_size

        print(f"  [TAR] Done. {len(results)} keyword matches.")
        return results

    except Exception as e:
        print(f"  [TAR] Error: {e}")
        import traceback
        traceback.print_exc()
        return [{'_error': f'Errore: {str(e)[:100]}'}]


# ─── Excel output ─────────────────────────────────────────────────────────────

BLUE_FONT = Font(color='0000FF', underline='single')
HEADER_FONT = Font(bold=True, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')

def write_sheet(ws, records: list[dict]):
    """Write results to an openpyxl worksheet."""
    headers = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']

    # Header row
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    if not records:
        ws.cell(row=2, column=1, value='Nessun risultato')
        return

    # Check for error
    if len(records) == 1 and '_error' in records[0]:
        err_msg = records[0]['_error']
        cell = ws.cell(row=2, column=1, value=f'⚠ {err_msg}')
        cell.font = Font(color='CC0000')
        return

    # Data rows
    for row_idx, rec in enumerate(records, start=2):
        data_val = rec.get('data', '')
        oggetto = rec.get('oggetto', '')
        descrizione = oggetto[:150]
        link = rec.get('link', '')

        ws.cell(row=row_idx, column=1, value=data_val)
        ws.cell(row=row_idx, column=2, value=oggetto)
        ws.cell(row=row_idx, column=3, value=descrizione)

        link_cell = ws.cell(row=row_idx, column=4)
        if link:
            link_cell.value = 'Apri'
            link_cell.hyperlink = link
            link_cell.font = BLUE_FONT
        else:
            link_cell.value = ''

        # Wrap oggetto
        ws.cell(row=row_idx, column=2).alignment = Alignment(wrap_text=True)

    # Column widths
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 80
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 12


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print(f"MONITORAGGIO ATTI CALABRIA - {TODAY.strftime('%d/%m/%Y')}")
    print("=" * 60)
    print(f"Periodo filtro ASP/Regione: {DATE_FROM_2DAYS} → {TODAY}")
    print(f"Periodo filtro TAR: {DATE_FROM_3DAYS} → {TODAY}")
    print()

    all_data = {}

    # ── 1. ASP portals ──────────────────────────────────────────────────────
    for portal_name, portal_url in ASP_PORTALS.items():
        results = scrape_asp_portal(portal_name, portal_url)
        all_data[portal_name] = results

    # ── 2. Regione Calabria ─────────────────────────────────────────────────
    all_data['Regione Calabria'] = scrape_regione_calabria()

    # ── 3. TAR Catanzaro ────────────────────────────────────────────────────
    all_data['TAR Catanzaro'] = scrape_tar_catanzaro()

    # ── 4. Write Excel ──────────────────────────────────────────────────────
    filename = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    output_path = f"/home/user/NS/{filename}"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    sheet_order = [
        'ASP Cosenza',
        'ASP Catanzaro',
        'ASP Crotone',
        'ASP Reggio Calabria',
        'ASP Vibo Valentia',
        'Regione Calabria',
        'TAR Catanzaro',
    ]

    for sheet_name in sheet_order:
        ws = wb.create_sheet(title=sheet_name)
        records = all_data.get(sheet_name, [])
        write_sheet(ws, records)
        count = len([r for r in records if '_error' not in r])
        print(f"  Sheet '{sheet_name}': {count} matches")

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")
    return output_path, all_data


if __name__ == '__main__':
    main()
