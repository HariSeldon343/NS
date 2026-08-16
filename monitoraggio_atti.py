#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Portali: 5 ASP, Regione Calabria, TAR Catanzaro
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import json
import warnings
from datetime import datetime, timedelta
import time

warnings.filterwarnings('ignore')

# ==================== DATE CONFIGURATION ====================
today = datetime.now()
DATE_FROM_2D = (today - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3D = (today - timedelta(days=3)).strftime('%Y-%m-%d')
DATE_TO = today.strftime('%Y-%m-%d')
FILE_DATE = today.strftime('%d-%m-%Y')

print(f"Esecuzione: {today.strftime('%d/%m/%Y %H:%M')}")
print(f"Filtro 2 giorni: dal {DATE_FROM_2D} al {DATE_TO}")
print(f"Filtro 3 giorni (TAR): dal {DATE_FROM_3D} al {DATE_TO}")

# ==================== KEYWORDS ====================
KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia',
    "Parere commissione", "Presa d'atto verifica", 'Programmazione',
    'Rete riabilitativa', 'Rete territoriale', 'Riabilitazione estensiva',
    'Riconversione prestazioni', 'Rinnovo accreditamento', 'San Teodoro',
    'Savelli Hospital', 'Starbene', 'Verifica requisiti',
    'Villa San Giuseppe', 'Villa del Rosario'
]
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    text_upper = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in text_upper:
            return True
    return bool(LIFE_RE.search(text))


# ==================== HTTP SESSION ====================
def make_session() -> requests.Session:
    s = requests.Session()
    s.verify = False
    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        ),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
    })
    return s


# ==================== ASP PORTALS ====================
ASP_PORTALS = [
    ('ASP Cosenza',       'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Catanzaro',     'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Crotone',       'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Reggio Calabria', 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Vibo Valentia', 'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
]


def _parse_asp_results_page(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Parse one results page from an ASP portal."""
    records = []
    table = soup.find('table')
    if not table:
        return records

    headers = []
    header_row = table.find('tr')
    if header_row:
        headers = [th.get_text(strip=True).lower() for th in header_row.find_all(['th', 'td'])]

    for row in table.find_all('tr')[1:]:
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        row_dict = {}
        for i, cell in enumerate(cells):
            key = headers[i] if i < len(headers) else f'col{i}'
            row_dict[key] = cell.get_text(strip=True)
            # Grab link if present
            link = cell.find('a', href=True)
            if link:
                href = link['href']
                if not href.startswith('http'):
                    domain = '/'.join(base_url.split('/')[:3])
                    href = domain + ('/' if not href.startswith('/') else '') + href
                row_dict[f'{key}_link'] = href

        # Identify oggetto field
        oggetto = (
            row_dict.get('oggetto') or
            row_dict.get('descrizione') or
            row_dict.get('titolo') or
            next((v for k, v in row_dict.items() if 'oggetto' in k or 'descr' in k), '') or
            ''
        )
        row_dict['_oggetto'] = oggetto

        # Identify date
        data = (
            row_dict.get('data pubblicazione') or
            row_dict.get('data') or
            row_dict.get('datapubblicazione') or
            next((v for k, v in row_dict.items() if 'data' in k), '') or
            ''
        )
        row_dict['_data'] = data

        # Identify link
        link_url = next(
            (v for k, v in row_dict.items() if k.endswith('_link')),
            base_url
        )
        row_dict['_link'] = link_url

        records.append(row_dict)

    return records


def _find_next_page_asp(soup: BeautifulSoup, base_url: str, current_page: int) -> str | None:
    """Find the next-page link in an ASP albo portal."""
    # Look for explicit "next" link
    for a in soup.find_all('a', href=True):
        text = a.get_text(strip=True).lower()
        href = a['href']
        if text in ['successivo', 'next', '>', '»', 'avanti'] or f'page={current_page + 1}' in href:
            if not href.startswith('http'):
                domain = '/'.join(base_url.split('/')[:3])
                href = domain + href
            return href

    # Look for pagination numbers
    for a in soup.find_all('a', href=True):
        href = a['href']
        text = a.get_text(strip=True)
        try:
            pg = int(text)
            if pg == current_page + 1:
                if not href.startswith('http'):
                    domain = '/'.join(base_url.split('/')[:3])
                    href = domain + href
                return href
        except ValueError:
            pass

    return None


def scrape_asp_portal(name: str, base_url: str) -> list[dict]:
    """Scrape one ASP albo online portal."""
    print(f"\n[{name}] Connessione a {base_url}...")
    results = []
    s = make_session()

    try:
        resp = s.get(base_url, timeout=12)
        resp.raise_for_status()
    except Exception as e:
        print(f"  ERRORE: {e}")
        return [{'_error': str(e)}]

    soup = BeautifulSoup(resp.text, 'html.parser')

    # Find and submit search form
    form = soup.find('form')
    form_action = base_url
    form_data = {}

    if form:
        form_action = form.get('action', base_url)
        if not form_action.startswith('http'):
            domain = '/'.join(base_url.split('/')[:3])
            form_action = domain + form_action
        for inp in form.find_all('input'):
            name_attr = inp.get('name')
            if name_attr:
                form_data[name_attr] = inp.get('value', '')

    # Set date filter (field id=dataPubblicazioneDal)
    form_data['dataPubblicazioneDal'] = DATE_FROM_2D
    form_data['dataPubblicazioneAl'] = DATE_TO

    try:
        resp = s.post(form_action, data=form_data, timeout=15,
                      headers={'Referer': base_url})
        resp.raise_for_status()
    except Exception as e:
        print(f"  ERRORE ricerca: {e}")
        return [{'_error': str(e)}]

    page = 1
    while True:
        soup = BeautifulSoup(resp.text, 'html.parser')
        rows = _parse_asp_results_page(soup, base_url)
        print(f"  Pag. {page}: {len(rows)} righe trovate")
        for row in rows:
            if matches_keywords(row.get('_oggetto', '')):
                results.append({
                    '_data': row.get('_data', ''),
                    '_oggetto': row.get('_oggetto', ''),
                    '_link': row.get('_link', base_url),
                })

        next_url = _find_next_page_asp(soup, base_url, page)
        if not next_url:
            break
        page += 1
        try:
            resp = s.get(next_url, timeout=15)
            resp.raise_for_status()
            time.sleep(0.3)
        except Exception as e:
            print(f"  ERRORE pag. {page}: {e}")
            break

    print(f"  Match: {len(results)}")
    return results


# ==================== REGIONE CALABRIA ====================
def scrape_regione_calabria() -> list[dict]:
    """Scrape provvedimenti della Regione Calabria."""
    print("\n[Regione Calabria] Avvio scraping...")
    results = []
    s = make_session()
    base = "https://www.regione.calabria.it/provvedimenti-della-regione/"

    page = 1
    while True:
        params = {
            'filter_date_accept_from': DATE_FROM_2D,
            'filter_date_accept_to': DATE_TO,
        }
        if page > 1:
            params['paged'] = str(page)

        try:
            resp = s.get(base, params=params, timeout=20)
            resp.raise_for_status()
        except Exception as e:
            print(f"  ERRORE pag. {page}: {e}")
            break

        soup = BeautifulSoup(resp.text, 'html.parser')

        # Count
        count_span = soup.find(string=lambda s: s and 'elementi utili' in str(s))
        if count_span and page == 1:
            print(f"  Totale risultati: {count_span.strip()}")

        table = soup.find('table')
        if not table:
            print(f"  Pag. {page}: nessuna tabella")
            break

        rows = table.find_all('tr')[1:]  # skip header
        if not rows:
            break

        print(f"  Pag. {page}: {len(rows)} righe")

        for row in rows:
            cells = row.find_all(['td', 'th'])
            if len(cells) < 5:
                continue
            # Columns: Tipologia, Data Repertoriazione, N, Dipartimento, Oggetto, Dettaglio
            tipologia = cells[0].get_text(strip=True)
            data_rep = cells[1].get_text(strip=True)
            numero = cells[2].get_text(strip=True)
            dipartimento = cells[3].get_text(strip=True)
            oggetto = cells[4].get_text(strip=True) if len(cells) > 4 else ''

            # Link from Dettaglio cell
            link = base
            if len(cells) > 5:
                a = cells[5].find('a', href=True)
                if a:
                    link = a['href']
                    if not link.startswith('http'):
                        link = 'https://www.regione.calabria.it' + link
            else:
                a = row.find('a', href=True)
                if a:
                    link = a['href']
                    if not link.startswith('http'):
                        link = 'https://www.regione.calabria.it' + link

            if matches_keywords(oggetto):
                results.append({
                    '_data': data_rep,
                    '_oggetto': oggetto,
                    '_link': link,
                })

        # Check for next page
        next_link = soup.find('a', string=lambda s: s and 'successiv' in s.lower())
        if not next_link:
            # Check numeric pagination
            paged_links = soup.find_all('a', href=re.compile(r'paged=\d+'))
            current_page_nums = [int(re.search(r'paged=(\d+)', a['href']).group(1))
                                  for a in paged_links if re.search(r'paged=(\d+)', a['href'])]
            if current_page_nums and page < max(current_page_nums):
                page += 1
                time.sleep(0.3)
                continue
            break
        page += 1
        time.sleep(0.3)

    print(f"  Match: {len(results)}")
    return results


# ==================== TAR CATANZARO ====================
TAR_AJAX_URL = (
    "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
    "?p_p_id=it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe"
    "&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view"
    "&p_p_resource_id=%2Fadministrative-acts%2Fsearch%2Fresults"
    "&p_p_cacheability=cacheLevelPage"
)
TAR_DOC_URL = (
    "https://mdp.giustizia-amministrativa.it/visualizza/"
    "?nodeRef=&schema=tar_cz&nrg={nrg}&nomeFile={nomeFile}&subDir=Provvedimenti"
)
TAR_FORM_URL_BASE = (
    "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
    "?p_p_id=it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe"
    "&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view"
    "&_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_javax.portlet.action=%2Fadministrative-acts%2Fsearch"
)
PREFIX = "_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"


def scrape_tar_catanzaro() -> list[dict]:
    """Scrape provvedimenti TAR Catanzaro."""
    print("\n[TAR Catanzaro] Avvio scraping...")
    results = []
    s = make_session()
    page_url = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"

    # Step 1: GET page to get session cookies and form details
    try:
        r = s.get(page_url, timeout=20)
        r.raise_for_status()
    except Exception as e:
        print(f"  ERRORE GET pagina: {e}")
        return [{'_error': str(e)}]

    soup = BeautifulSoup(r.text, 'html.parser')
    form = None
    for f in soup.find_all('form'):
        if 'jjYpzZYF4Qfe' in f.get('action', '') and 'search' in f.get('action', ''):
            form = f
            break

    if not form:
        print("  ERRORE: form di ricerca non trovato")
        return [{'_error': 'Form non trovato'}]

    action = form.get('action')
    formdate_inp = form.find('input', attrs={'name': f'{PREFIX}formDate'})
    formdate = formdate_inp['value'] if formdate_inp else str(int(datetime.now().timestamp() * 1000))

    # Step 2: POST the search form to initialize the session
    form_data = {
        f'{PREFIX}formDate': formdate,
        f'{PREFIX}publishDateFrom': DATE_FROM_3D,
        f'{PREFIX}publishDateTo': DATE_TO,
        f'{PREFIX}year': '', f'{PREFIX}number': '',
        f'{PREFIX}hearingDateFrom': '', f'{PREFIX}hearingDateTo': '',
        f'{PREFIX}section': '', f'{PREFIX}type': '',
        f'{PREFIX}specific': '', f'{PREFIX}president': '',
        f'{PREFIX}draftingJudge': '',
    }
    try:
        r2 = s.post(action, data=form_data, timeout=30,
                    headers={'Referer': page_url,
                             'Origin': 'https://www.giustizia-amministrativa.it'})
        r2.raise_for_status()
    except Exception as e:
        print(f"  ERRORE POST form: {e}")
        return [{'_error': str(e)}]

    # Extract updated additionalInfo from the response page script
    additional_info = {
        "schema": "TAR_CATANZARO", "type": None,
        "year": "", "number": "",
        "hearingDateFrom": None, "hearingDateTo": None,
        "publishDateFrom": DATE_FROM_3D, "publishDateTo": DATE_TO,
        "hearingType": None, "nrg": None, "section": "",
        "provisionSpecification": "", "president": "", "draftingJudge": "",
        "subjectMatter": None, "page": None, "size": None,
        "orderBy": None, "orderStrategy": None, "queryString": None
    }

    # Retrieve all records in chunks (DataTables server-side)
    start = 0
    chunk = 500
    total = None

    while True:
        payload = {
            "draw": (start // chunk) + 1,
            "columns": [
                {"data": c, "name": "", "searchable": True, "orderable": True,
                 "search": {"value": "", "regex": False}}
                for c in ["nrgFascicolo", "sezione", "parte", "tipoUdienza",
                           "dataUdienza", "numProvvedimento", "dataPubblicazione",
                           "tipoProvvedimento", "relatore", "presidente", "esito"]
            ],
            "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
            "start": start,
            "length": chunk,
            "search": {"value": "", "regex": False},
            "additionalInfo": json.dumps(additional_info)
        }

        try:
            r3 = s.post(TAR_AJAX_URL, json=payload, timeout=30,
                        headers={
                            'Referer': r2.url,
                            'Origin': 'https://www.giustizia-amministrativa.it',
                            'Accept': 'application/json, text/javascript, */*; q=0.01',
                            'Content-Type': 'application/json',
                            'X-Requested-With': 'XMLHttpRequest',
                        })
            r3.raise_for_status()
            data_json = r3.json()
        except Exception as e:
            print(f"  ERRORE AJAX start={start}: {e}")
            break

        records = data_json.get('data', [])
        if total is None:
            total = data_json.get('recordsTotal', 0)
            print(f"  Totale provvedimenti nel periodo: {total}")

        print(f"  Chunk start={start}: {len(records)} record")

        for rec in records:
            parte = rec.get('parte', '')
            nrg = rec.get('nrgFascicolo', '')
            nome_file = rec.get('nomeFile', '')
            data_pub = rec.get('dataPubblicazione', '')
            tipo = rec.get('tipoProvvedimento', '')

            # Build document URL
            doc_url = TAR_DOC_URL.format(
                nrg=requests.utils.quote(nrg),
                nomeFile=requests.utils.quote(nome_file)
            ) if nome_file else page_url

            # For TAR the "oggetto" column is "parte"
            if matches_keywords(parte):
                results.append({
                    '_data': data_pub,
                    '_oggetto': f"{parte} — {tipo}",
                    '_link': doc_url,
                })

        start += len(records)
        if start >= total or not records:
            break
        time.sleep(0.2)

    print(f"  Match: {len(results)}")
    return results


# ==================== EXCEL OUTPUT ====================
LINK_FONT = Font(color='0563C1', underline='single')
HEADER_FONT = Font(bold=True, color='FFFFFF')
HEADER_FILL = PatternFill(fill_type='solid', fgColor='2E4A7F')
ALT_FILL = PatternFill(fill_type='solid', fgColor='EEF2F7')
HEADERS = ['Data', 'Oggetto', 'Descrizione breve', 'Link']


def _write_sheet(ws, portal_name: str, rows: list[dict], error_msg: str = ''):
    """Write data rows to an Excel sheet."""
    ws.title = portal_name[:31]

    # Header row
    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    if error_msg:
        cell = ws.cell(row=2, column=1, value=f"Portale non raggiungibile: {error_msg}")
        cell.font = Font(italic=True, color='CC0000')
        ws.merge_cells('A2:D2')
        _autofit(ws, [[error_msg, '', '', '']])
        return

    if not rows:
        cell = ws.cell(row=2, column=1, value='Nessun risultato')
        cell.font = Font(italic=True, color='888888')
        ws.merge_cells('A2:D2')
        _autofit(ws, [['Nessun risultato', '', '', '']])
        return

    for row_idx, row in enumerate(rows, 2):
        oggetto = row.get('_oggetto', '')
        descrizione = oggetto[:150]
        data = row.get('_data', '')
        link = row.get('_link', '')

        fill = ALT_FILL if row_idx % 2 == 0 else None

        c1 = ws.cell(row=row_idx, column=1, value=data)
        c2 = ws.cell(row=row_idx, column=2, value=oggetto)
        c3 = ws.cell(row=row_idx, column=3, value=descrizione)
        c4 = ws.cell(row=row_idx, column=4, value='Apri atto')

        if fill:
            for c in [c1, c2, c3, c4]:
                c.fill = fill

        c2.alignment = Alignment(wrap_text=True)

        if link:
            c4.hyperlink = link
            c4.font = LINK_FONT
        else:
            c4.value = '—'

    _autofit(ws, rows)


def _autofit(ws, rows):
    """Set reasonable column widths."""
    widths = [18, 70, 55, 15]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22


def generate_xlsx(all_data: dict[str, list[dict]], output_path: str):
    """Generate the final Excel file."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    SHEET_ORDER = [
        'ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
        'ASP Reggio Calabria', 'ASP Vibo Valentia',
        'Regione Calabria', 'TAR Catanzaro'
    ]

    for sheet_name in SHEET_ORDER:
        ws = wb.create_sheet()
        data = all_data.get(sheet_name, [])

        # Check if it's an error record
        if data and len(data) == 1 and '_error' in data[0]:
            _write_sheet(ws, sheet_name, [], error_msg=data[0]['_error'])
        else:
            _write_sheet(ws, sheet_name, data)

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")


# ==================== MAIN ====================
def main():
    all_data = {}

    # --- ASP Portals ---
    for name, url in ASP_PORTALS:
        rows = scrape_asp_portal(name, url)
        if rows and '_error' in rows[0]:
            all_data[name] = rows  # error record
        else:
            all_data[name] = [r for r in rows if '_error' not in r]

    # --- Regione Calabria ---
    all_data['Regione Calabria'] = scrape_regione_calabria()

    # --- TAR Catanzaro ---
    all_data['TAR Catanzaro'] = scrape_tar_catanzaro()

    # --- Summary ---
    print("\n" + "="*60)
    print("RIEPILOGO MATCH")
    print("="*60)
    for name, data in all_data.items():
        if data and '_error' in data[0]:
            print(f"  {name}: ERRORE ({data[0]['_error'][:60]})")
        else:
            print(f"  {name}: {len(data)} match")

    # --- Generate Excel ---
    output_path = f"/home/user/NS/Monitoraggio_Atti_Calabria_{FILE_DATE}.xlsx"
    generate_xlsx(all_data, output_path)
    return output_path


if __name__ == '__main__':
    output = main()
    print(f"\nOutput: {output}")
