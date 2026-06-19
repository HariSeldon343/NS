#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria - Script di raccolta automatica
Portali: ASP Cosenza, Catanzaro, Crotone, Reggio Calabria, Vibo Valentia,
         Regione Calabria, TAR Catanzaro
"""

import requests
from bs4 import BeautifulSoup
from datetime import date, timedelta
import re
import time
import warnings
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

TODAY = date.today()
DATE_FROM_2D = (TODAY - timedelta(days=2)).isoformat()
DATE_FROM_3D = (TODAY - timedelta(days=3)).isoformat()
DATE_TO = TODAY.isoformat()
NOW_STR = TODAY.strftime('%d/%m/%Y')
RUN_TIMESTAMP = TODAY.strftime('%d-%m-%Y')

KEYWORDS = [
    (r'\bADI\b',                          'ADI'),
    (r'Assistenza domiciliare',            'Assistenza domiciliare'),
    (r'\bANMIC\b',                         'ANMIC'),
    (r'Accreditamento',                    'Accreditamento'),
    (r'Aumento di budget',                 'Aumento di budget'),
    (r'Autismo',                           'Autismo'),
    (r'Autorizzazione all[^\w]esercizio',  'Autorizzazione all\'esercizio'),
    (r'Autorizzazione alla realizzazione', 'Autorizzazione alla realizzazione'),
    (r'Autorizzazioni',                    'Autorizzazioni'),
    (r'\bBudget\b',                        'Budget'),
    (r'Casa Giardino',                     'Casa Giardino'),
    (r'Centro San Giuseppe',               'Centro San Giuseppe'),
    (r'Centro salute e benessere',         'Centro salute e benessere'),
    (r'Fabbisogni LEA',                    'Fabbisogni LEA'),
    (r'Fisiolab',                          'Fisiolab'),
    (r'Fisioterapia',                      'Fisioterapia'),
    (r'\bLIFE\b',                          'LIFE'),
    (r'Parere commissione',                'Parere commissione'),
    (r'Presa d[^\w]atto verifica',         'Presa d\'atto verifica'),
    (r'Programmazione',                    'Programmazione'),
    (r'Rete riabilitativa',                'Rete riabilitativa'),
    (r'Rete territoriale',                 'Rete territoriale'),
    (r'Riabilitazione estensiva',          'Riabilitazione estensiva'),
    (r'Riconversione prestazioni',         'Riconversione prestazioni'),
    (r'Rinnovo accreditamento',            'Rinnovo accreditamento'),
    (r'San Teodoro',                       'San Teodoro'),
    (r'Savelli Hospital',                  'Savelli Hospital'),
    (r'Starbene',                          'Starbene'),
    (r'Verifica requisiti',                'Verifica requisiti'),
    (r'Villa San Giuseppe',                'Villa San Giuseppe'),
    (r'Villa del Rosario',                 'Villa del Rosario'),
]

COMPILED_KEYWORDS = [(re.compile(pat, re.IGNORECASE), label) for pat, label in KEYWORDS]


def matches_keywords(text):
    """Returns (matched, keyword_label) or (False, None)."""
    if not text:
        return False, None
    for pattern, label in COMPILED_KEYWORDS:
        if pattern.search(text):
            return True, label
    return False, None


def make_session():
    s = requests.Session()
    s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8',
    })
    return s


# ===========================================================================
# ASP Portals (SISR)
# ===========================================================================

ASP_PORTALS = {
    'ASP Cosenza':         'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':       'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':         'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria': 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':   'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}


def scrape_asp_portal(portal_name, base_url):
    print(f"\n[{portal_name}] Scraping: {base_url}")
    results = []
    s = make_session()

    try:
        r0 = s.get(base_url, timeout=25, verify=False)

        if r0.status_code == 503:
            msg = (f"Portale non disponibile (HTTP 503 - Backend SISR non raggiungibile).\n"
                   f"I server SISR ({base_url}) risultano temporaneamente non raggiungibili "
                   f"dall'ambiente cloud. Verificare manualmente il portale o riprovare in orario lavorativo.")
            print(f"  [503] {msg[:80]}")
            return [], msg

        r0.raise_for_status()
        soup0 = BeautifulSoup(r0.text, 'lxml')

        form = soup0.find('form')
        if not form:
            return [], "Form di ricerca non trovato nella pagina"

        form_action = form.get('action', base_url)
        if not form_action.startswith('http'):
            from urllib.parse import urljoin
            form_action = urljoin(base_url, form_action)
        form_method = form.get('method', 'get').lower()

        # Collect form base fields
        base_data = {}
        for inp in form.find_all(['input', 'select']):
            name = inp.get('name')
            if not name:
                continue
            val = inp.get('value', '')
            if inp.name == 'select':
                opt = inp.find('option', selected=True) or inp.find('option')
                val = opt.get('value', '') if opt else ''
            inp_type = inp.get('type', '').lower()
            if inp_type not in ('submit', 'button', 'reset', 'image'):
                base_data[name] = val

        # Set date filter
        for date_field in ['dataPubblicazioneDal', 'dataDal', 'dataInizio', 'dataFrom']:
            if date_field in base_data or form.find(id=date_field):
                base_data[date_field] = DATE_FROM_2D
                break
        else:
            base_data['dataPubblicazioneDal'] = DATE_FROM_2D

        for date_field in ['dataPubblicazioneAl', 'dataAl', 'dataFine', 'dataTo']:
            if date_field in base_data or form.find(id=date_field):
                base_data[date_field] = DATE_TO
                break
        else:
            base_data['dataPubblicazioneAl'] = DATE_TO

        print(f"  Invio ricerca: {DATE_FROM_2D} → {DATE_TO}")

        page = 1
        max_pages = 50

        while page <= max_pages:
            page_data = dict(base_data)
            if page > 1:
                for pk in ['page', 'pagina', 'currentPage', 'p']:
                    page_data[pk] = page

            try:
                if form_method == 'post':
                    r = s.post(form_action, data=page_data, timeout=25, verify=False)
                else:
                    r = s.get(form_action, params=page_data, timeout=25, verify=False)
                r.raise_for_status()
            except requests.exceptions.Timeout:
                print(f"  [TIMEOUT] Pagina {page}")
                break
            except requests.exceptions.RequestException as e:
                print(f"  [ERROR] Pagina {page}: {e}")
                break

            soup = BeautifulSoup(r.text, 'lxml')
            tables = soup.find_all('table')
            if not tables:
                print(f"  Pagina {page}: nessuna tabella")
                break

            page_rows = 0
            for table in tables:
                rows = table.find_all('tr')
                for row in rows[1:]:
                    cells = row.find_all('td')
                    if len(cells) < 2:
                        continue

                    cell_texts = [c.get_text(separator=' ', strip=True) for c in cells]
                    full_text = ' '.join(cell_texts)

                    data_cell = next((t for t in cell_texts if re.match(r'\d{2}[/\-]\d{2}[/\-]\d{4}', t)), '')
                    oggetto_cell = max(cell_texts, key=len, default='')

                    link_url = ''
                    a = row.find('a', href=True)
                    if a:
                        href = a.get('href', '')
                        if not href.startswith('http'):
                            from urllib.parse import urljoin
                            href = urljoin(base_url, href)
                        link_url = href

                    matched, kw = matches_keywords(full_text)
                    if matched:
                        results.append({
                            'data': data_cell,
                            'oggetto': oggetto_cell,
                            'keyword': kw,
                            'link': link_url,
                        })
                        print(f"  [MATCH:{kw}] {oggetto_cell[:80]}")
                    page_rows += 1

            print(f"  Pagina {page}: {page_rows} righe, {len(results)} match totali")
            if page_rows == 0:
                break

            next_link = soup.find('a', string=re.compile(r'successiv|next|>', re.IGNORECASE))
            if not next_link:
                pag_links = [a for a in soup.find_all('a', href=True, string=re.compile(r'^\d+$'))]
                page_nums = [int(a.get_text(strip=True)) for a in pag_links if a.get_text(strip=True).isdigit()]
                if page_nums and page < max(page_nums):
                    page += 1
                else:
                    break
            else:
                page += 1
            time.sleep(0.5)

        return results, None

    except requests.exceptions.Timeout:
        msg = (f"Timeout: portale SISR non raggiungibile.\n"
               f"I portali {base_url} sono accessibili solo da reti autorizzate "
               f"(VPN regionale o IP italiani). Riprovare eseguendo lo script localmente.")
        print(f"  [TIMEOUT] {msg[:80]}")
        return [], msg
    except requests.exceptions.ConnectionError as e:
        msg = f"Connessione rifiutata: {str(e)[:200]}"
        print(f"  [CONN ERROR] {msg[:80]}")
        return [], msg
    except Exception as e:
        msg = f"Errore imprevisto: {str(e)[:300]}"
        print(f"  [ERROR] {msg[:80]}")
        return [], msg


# ===========================================================================
# Regione Calabria
# ===========================================================================

def scrape_regione_calabria():
    print("\n[Regione Calabria] Scraping provvedimenti...")
    results = []
    s = make_session()
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'

    page = 1
    max_pages = 100

    while page <= max_pages:
        params = {
            'filter_date_from': DATE_FROM_2D,
            'filter_date_to': DATE_TO,
            'filter_active': 'true',
            'sort_order': 'DESC',
        }
        if page > 1:
            params['paged'] = page

        try:
            r = s.get(base_url, params=params, timeout=30)
            r.raise_for_status()
        except Exception as e:
            print(f"  Errore pagina {page}: {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table', class_='table')

        if not table:
            print(f"  Pagina {page}: nessuna tabella")
            break

        rows = table.find_all('tr')[1:]
        if not rows:
            break

        page_matches = 0
        for row in rows:
            cells = row.find_all(['td', 'th'])
            if len(cells) < 5:
                continue

            tipologia = cells[0].get_text(strip=True)
            data_cell = cells[1].get_text(strip=True)
            numero = cells[2].get_text(strip=True)
            dipartimento = cells[3].get_text(strip=True)
            oggetto = cells[4].get_text(strip=True)

            # Link from last column
            link_url = ''
            for c in reversed(cells):
                a = c.find('a', href=True)
                if a:
                    link_url = a['href']
                    break

            search_text = f"{tipologia} {oggetto} {dipartimento}"
            matched, kw = matches_keywords(search_text)

            if matched:
                oggetto_full = f"[{tipologia} N.{numero}] {oggetto}" if tipologia else oggetto
                results.append({
                    'data': data_cell,
                    'oggetto': oggetto_full,
                    'keyword': kw,
                    'link': link_url,
                })
                page_matches += 1
                print(f"  [MATCH:{kw}] {data_cell} | {oggetto[:80]}")

        print(f"  Pagina {page}: {len(rows)} righe, {page_matches} match")

        # Check next page
        has_next = bool(soup.find('a', href=lambda h: h and f'paged={page+1}' in h))
        if not has_next:
            break
        page += 1
        time.sleep(0.3)

    print(f"  Totale: {len(results)} atti corrispondenti")
    return results, None


# ===========================================================================
# TAR Catanzaro
# ===========================================================================

def scrape_tar_catanzaro():
    print("\n[TAR Catanzaro] Scraping provvedimenti...")
    results = []
    s = make_session()
    base_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'

    try:
        r0 = s.get(base_url, timeout=20)
        r0.raise_for_status()
        soup0 = BeautifulSoup(r0.text, 'lxml')

        form = soup0.find('form', id=lambda x: x and 'appeal-fm' in str(x))
        if not form:
            return [], "Form di ricerca non trovato nella pagina TAR Catanzaro"

        action = form.get('action', '')
        form_id = form.get('id', '')
        ns_match = re.search(r'(_\w+_INSTANCE_\w+_)', form_id) or re.search(r'(_\w+_INSTANCE_\w+_)', action)
        ns = ns_match.group(1) if ns_match else '_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_'

        formDate_input = form.find('input', type='hidden')
        formDate = formDate_input.get('value', '') if formDate_input else ''

        print(f"  Ricerca dal: {DATE_FROM_3D} al: {DATE_TO}")

        r1 = s.post(action, data={
            f'{ns}formDate': formDate,
            f'{ns}publishDateFrom': DATE_FROM_3D,
            f'{ns}publishDateTo': DATE_TO,
        }, headers={
            'Content-Type': 'application/x-www-form-urlencoded',
            'Referer': r0.url,
            'Origin': 'https://www.giustizia-amministrativa.it',
        }, timeout=30)
        r1.raise_for_status()

        soup1 = BeautifulSoup(r1.text, 'lxml')

        # Find portlet content
        portlet = soup1.find('div', id=lambda x: x and 'INSTANCE' in str(x) and 'Administrative' in str(x))
        if not portlet:
            portlet = soup1

        raw_portlet = str(portlet)

        # Check for Java NullPointerException
        if '<!-- [ERROR' in raw_portlet or 'NullPointerException' in raw_portlet:
            msg = (
                "Il portale TAR Catanzaro (giustizia-amministrativa.it) restituisce un errore "
                "lato server (Java NullPointerException) per ogni ricerca effettuata da ambienti "
                "cloud/automatizzati. Il backend Liferay sembra richiedere un contesto browser "
                "interattivo. Per verificare manualmente: https://www.giustizia-amministrativa.it/"
                "provvedimenti-tar-catanzaro - impostare 'Data Pubblicazione dal' a "
                f"{DATE_FROM_3D} e cercare."
            )
            print(f"  [ERROR] Java NPE - backend non accessibile da cloud")
            return [], msg

        # Parse results tables
        tables = portlet.find_all('table')
        print(f"  Tabelle trovate: {len(tables)}")

        for table in tables:
            rows = table.find_all('tr')[1:]
            for row in rows:
                cells = row.find_all('td')
                if not cells:
                    continue

                cell_texts = [c.get_text(separator=' ', strip=True) for c in cells]
                full_text = ' '.join(cell_texts)

                data_cell = next((t for t in cell_texts if re.match(r'\d{2}[/\-]\d{2}[/\-]\d{4}', t)), '')
                link_url = ''
                a = row.find('a', href=True)
                if a:
                    href = a['href']
                    if not href.startswith('http'):
                        href = 'https://www.giustizia-amministrativa.it' + href
                    link_url = href

                matched, kw = matches_keywords(full_text)
                if matched:
                    results.append({
                        'data': data_cell,
                        'oggetto': full_text[:500],
                        'keyword': kw,
                        'link': link_url,
                    })
                    print(f"  [MATCH:{kw}] {full_text[:80]}")

        print(f"  Totale: {len(results)} atti corrispondenti")
        return results, None

    except requests.exceptions.Timeout:
        return [], f"Timeout connessione al portale TAR Catanzaro"
    except Exception as e:
        return [], f"Errore: {str(e)[:300]}"


# ===========================================================================
# Excel Generation
# ===========================================================================

BLUE_DARK    = '1F4E79'
BLUE_LIGHT   = 'D6E4F0'
YELLOW_WARN  = 'FFF2CC'
RED_WARN     = 'FFE0E0'
GREEN_OK     = 'E2EFDA'
GREY_TEXT    = '666666'
LINK_COLOR   = '0563C1'
WHITE        = 'FFFFFF'


def _fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type='solid')


def _border():
    thin = Side(style='thin', color='CCCCCC')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def build_excel(all_data, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Summary sheet ----
    ws_sum = wb.create_sheet(title='Riepilogo')
    ws_sum.column_dimensions['A'].width = 25
    ws_sum.column_dimensions['B'].width = 12
    ws_sum.column_dimensions['C'].width = 60

    ws_sum.cell(1, 1, f'MONITORAGGIO ATTI CALABRIA – {NOW_STR}').font = Font(bold=True, size=14, color=BLUE_DARK)
    ws_sum.merge_cells('A1:C1')
    ws_sum.row_dimensions[1].height = 30

    ws_sum.cell(2, 1, f'Periodo ASP/Regione: {DATE_FROM_2D} → {DATE_TO}').font = Font(size=10, italic=True)
    ws_sum.merge_cells('A2:C2')
    ws_sum.cell(3, 1, f'Periodo TAR: {DATE_FROM_3D} → {DATE_TO}').font = Font(size=10, italic=True)
    ws_sum.merge_cells('A3:C3')
    ws_sum.cell(4, 1, '').value = None

    hdr = ['Portale', 'Atti trovati', 'Stato']
    for col, h in enumerate(hdr, 1):
        c = ws_sum.cell(5, col, h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = _fill(BLUE_DARK)
        c.alignment = Alignment(horizontal='center')
        c.border = _border()

    sheet_order = [
        'ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
        'ASP Reggio Calabria', 'ASP Vibo Valentia',
        'Regione Calabria', 'TAR Catanzaro',
    ]

    for i, name in enumerate(sheet_order, 6):
        info = all_data.get(name, {})
        n = len(info.get('results', []))
        err = info.get('error')

        stato_text = 'OK' if not err else 'ERRORE'
        if not err and n == 0:
            stato_text = 'Nessun match'

        fill_color = GREEN_OK if not err and n > 0 else (YELLOW_WARN if not err else RED_WARN)

        ws_sum.cell(i, 1, name).font = Font(size=10, bold=True)
        ws_sum.cell(i, 1).border = _border()

        c_num = ws_sum.cell(i, 2, n if not err else '—')
        c_num.alignment = Alignment(horizontal='center')
        c_num.font = Font(size=10, bold=(n > 0))
        c_num.fill = _fill(fill_color)
        c_num.border = _border()

        c_stato = ws_sum.cell(i, 3, stato_text if not err else f'ERRORE - {err[:80]}')
        c_stato.font = Font(size=9, italic=bool(err))
        c_stato.fill = _fill(fill_color)
        c_stato.alignment = Alignment(wrap_text=True)
        c_stato.border = _border()
        ws_sum.row_dimensions[i].height = 30

    # ---- Data sheets ----
    for sheet_name in sheet_order:
        ws = wb.create_sheet(title=sheet_name)

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 65
        ws.column_dimensions['C'].width = 38
        ws.column_dimensions['D'].width = 55

        # Header
        hdr_cols = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link all\'atto']
        for col, h in enumerate(hdr_cols, 1):
            c = ws.cell(1, col, h)
            c.font = Font(color=WHITE, bold=True, size=11)
            c.fill = _fill(BLUE_DARK)
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = _border()
        ws.row_dimensions[1].height = 28
        ws.freeze_panes = 'A2'

        info = all_data.get(sheet_name, {})
        results = info.get('results', [])
        error = info.get('error')

        if error and not results:
            ws.merge_cells('A2:D2')
            c = ws.cell(2, 1)
            c.value = f"PORTALE NON ACCESSIBILE in data {NOW_STR}"
            c.font = Font(bold=True, color='C00000', size=10)
            c.fill = _fill(RED_WARN)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = _border()
            ws.row_dimensions[2].height = 22

            ws.merge_cells('A3:D3')
            c3 = ws.cell(3, 1)
            c3.value = error
            c3.font = Font(size=9, color='555555', italic=True)
            c3.fill = _fill(YELLOW_WARN)
            c3.alignment = Alignment(wrap_text=True, vertical='top')
            c3.border = _border()
            ws.row_dimensions[3].height = 70
            continue

        if not results:
            ws.merge_cells('A2:D2')
            c = ws.cell(2, 1, f'Nessun atto trovato nel periodo {DATE_FROM_2D} → {DATE_TO} con le keyword specificate.')
            c.font = Font(italic=True, color=GREY_TEXT, size=10)
            c.fill = _fill(BLUE_LIGHT)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = _border()
            ws.row_dimensions[2].height = 22
            continue

        for idx, item in enumerate(results, 2):
            data_val  = item.get('data', '')
            oggetto   = item.get('oggetto', '')
            desc      = oggetto[:150]
            link      = item.get('link', '')
            kw        = item.get('keyword', '')

            row_fill = _fill(BLUE_LIGHT) if idx % 2 == 0 else PatternFill()

            c_d = ws.cell(idx, 1, data_val)
            c_d.font = Font(size=10)
            c_d.fill = row_fill
            c_d.alignment = Alignment(horizontal='center', vertical='top')
            c_d.border = _border()

            c_o = ws.cell(idx, 2, oggetto)
            c_o.font = Font(size=10)
            c_o.fill = row_fill
            c_o.alignment = Alignment(wrap_text=True, vertical='top')
            c_o.border = _border()

            c_b = ws.cell(idx, 3, desc)
            c_b.font = Font(size=9)
            c_b.fill = row_fill
            c_b.alignment = Alignment(wrap_text=True, vertical='top')
            c_b.border = _border()

            if link:
                c_l = ws.cell(idx, 4, link)
                c_l.hyperlink = link
                c_l.font = Font(color=LINK_COLOR, underline='single', size=9)
            else:
                c_l = ws.cell(idx, 4, 'N/D')
                c_l.font = Font(size=9, color=GREY_TEXT, italic=True)
            c_l.fill = row_fill
            c_l.alignment = Alignment(vertical='top', wrap_text=True)
            c_l.border = _border()

            ws.row_dimensions[idx].height = 50

        # Total row
        tot_row = len(results) + 2
        ws.merge_cells(f'A{tot_row}:D{tot_row}')
        c_tot = ws.cell(tot_row, 1, f'Totale atti trovati: {len(results)}  |  Periodo: {DATE_FROM_2D} → {DATE_TO}')
        c_tot.font = Font(bold=True, size=10, color=WHITE)
        c_tot.fill = _fill(BLUE_DARK)
        c_tot.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[tot_row].height = 22

        print(f"  Sheet '{sheet_name}': {len(results)} righe")

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    print("=" * 70)
    print(f"MONITORAGGIO ATTI CALABRIA - {NOW_STR}")
    print(f"Periodo ASP/Regione: {DATE_FROM_2D} → {DATE_TO}")
    print(f"Periodo TAR: {DATE_FROM_3D} → {DATE_TO}")
    print("=" * 70)

    all_data = {}

    for name, url in ASP_PORTALS.items():
        results, error = scrape_asp_portal(name, url)
        all_data[name] = {'results': results, 'error': error}

    rc_results, rc_error = scrape_regione_calabria()
    all_data['Regione Calabria'] = {'results': rc_results, 'error': rc_error}

    tar_results, tar_error = scrape_tar_catanzaro()
    all_data['TAR Catanzaro'] = {'results': tar_results, 'error': tar_error}

    filename = f"Monitoraggio_Atti_Calabria_{RUN_TIMESTAMP}.xlsx"
    output_path = f"/home/user/NS/{filename}"
    build_excel(all_data, output_path)

    # Summary
    print("\n" + "=" * 70)
    print("RIEPILOGO FINALE:")
    total_matches = 0
    for name in ['ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
                 'ASP Reggio Calabria', 'ASP Vibo Valentia',
                 'Regione Calabria', 'TAR Catanzaro']:
        info = all_data.get(name, {})
        n = len(info.get('results', []))
        e = info.get('error')
        total_matches += n
        status = f"{n} match" if not e else f"ERRORE ({str(e)[:50]})"
        print(f"  {name:<25} {status}")

    print(f"\nMatch totali: {total_matches}")
    print(f"File: {output_path}")
    return output_path, all_data, total_matches


if __name__ == '__main__':
    main()
