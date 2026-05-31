#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scarica e filtra atti da portali ASP Calabria, Regione Calabria e TAR Catanzaro
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import re
import time
import traceback
import urllib.parse
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# CONFIGURAZIONE DATE
# ============================================================
TODAY = datetime(2026, 5, 31)
DATE_2DAYS_AGO = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_3DAYS_AGO = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
FILE_DATE = TODAY.strftime('%d-%m-%Y')
OUTPUT_FILE = f'/home/user/NS/Monitoraggio_Atti_Calabria_{FILE_DATE}.xlsx'

print(f"Data esecuzione:          {TODAY.strftime('%d/%m/%Y')}")
print(f"Periodo ASP/Regione dal:  {DATE_2DAYS_AGO}")
print(f"Periodo TAR dal:          {DATE_3DAYS_AGO}")
print(f"Output:                   {OUTPUT_FILE}")
print("=" * 60)

# ============================================================
# KEYWORDS
# ============================================================
KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia', 'Parere commissione',
    "Presa d'atto verifica", 'Programmazione', 'Rete riabilitativa',
    'Rete territoriale', 'Riabilitazione estensiva', 'Riconversione prestazioni',
    'Rinnovo accreditamento', 'San Teodoro', 'Savelli Hospital', 'Starbene',
    'Verifica requisiti', 'Villa San Giuseppe', 'Villa del Rosario'
]

LIFE_PATTERN = re.compile(r'\bLIFE\b', re.IGNORECASE)


def matches_keywords(text):
    if not text:
        return False
    if LIFE_PATTERN.search(text):
        return True
    text_upper = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in text_upper:
            return True
    return False


# ============================================================
# HTTP SESSION
# ============================================================
session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,'
              'image/webp,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
})

PORTAL_ERRORS = {}  # name -> error message


def safe_get(url, verify=True, **kwargs):
    for attempt in range(3):
        try:
            r = session.get(url, timeout=25, verify=verify, **kwargs)
            r.raise_for_status()
            return r
        except requests.exceptions.Timeout:
            msg = f"Timeout al tentativo {attempt+1}/3"
            print(f"  {msg} per {url}")
            if attempt < 2:
                time.sleep(2 ** attempt)
        except requests.exceptions.HTTPError as e:
            print(f"  HTTP {e.response.status_code} al tentativo {attempt+1}/3 per {url}")
            if e.response.status_code in (503, 502, 504):
                if attempt < 2:
                    time.sleep(2 ** attempt)
            else:
                return None
        except Exception as e:
            print(f"  Errore al tentativo {attempt+1}/3: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None


def safe_post(url, data=None, verify=True, **kwargs):
    for attempt in range(3):
        try:
            r = session.post(url, data=data, timeout=25, verify=verify, **kwargs)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  POST attempt {attempt+1}/3 failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None


def make_absolute_url(href, base_url):
    if not href:
        return ''
    if href.startswith('http'):
        return href
    parsed = urllib.parse.urlparse(base_url)
    if href.startswith('//'):
        return f"{parsed.scheme}:{href}"
    if href.startswith('/'):
        return f"{parsed.scheme}://{parsed.netloc}{href}"
    base_path = parsed.path.rsplit('/', 1)[0]
    return f"{parsed.scheme}://{parsed.netloc}{base_path}/{href}"


# ============================================================
# SCRAPER ASP PORTALS (SISR AlboOnline)
# ============================================================

ASP_PORTALS = {
    'ASP Cosenza':        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria':'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}


def scrape_asp_portal(name, base_url, date_from):
    results = []
    print(f"\n[{name}] Connessione a {base_url}")

    try:
        resp = safe_get(base_url)
        if not resp:
            PORTAL_ERRORS[name] = (
                "Portale non raggiungibile (503 Service Unavailable / timeout). "
                "Il server SISR regione.calabria.it non risponde dalla rete cloud. "
                "Eseguire lo script in locale per accedere al portale."
            )
            print(f"[{name}] Portale non raggiungibile")
            return results

        soup = BeautifulSoup(resp.text, 'lxml')

        # Form action
        form = soup.find('form')
        if form and form.get('action'):
            form_action = make_absolute_url(form['action'], base_url)
        else:
            form_action = base_url
        print(f"[{name}] Form action: {form_action}")

        # Hidden fields + all non-submit inputs
        base_form_data = {}
        if form:
            for inp in form.find_all('input'):
                inp_name = inp.get('name', '')
                inp_type = inp.get('type', 'text').lower()
                if inp_name and inp_type not in ('submit', 'button', 'image', 'reset'):
                    base_form_data[inp_name] = inp.get('value', '')
            for sel in form.find_all('select'):
                sel_name = sel.get('name', '')
                if sel_name:
                    selected = sel.find('option', selected=True) or sel.find('option')
                    base_form_data[sel_name] = selected.get('value', '') if selected else ''

        base_form_data['dataPubblicazioneDal'] = date_from
        print(f"[{name}] Campi form: {list(base_form_data.keys())}")

        page = 1
        while True:
            form_data = base_form_data.copy()
            if page > 1:
                form_data['page'] = str(page)

            print(f"[{name}] Pagina {page}...")
            resp = safe_post(form_action, data=form_data)
            if not resp:
                break

            soup = BeautifulSoup(resp.text, 'lxml')
            if page == 1:
                print(f"[{name}] Titolo: {soup.title.string if soup.title else 'N/A'}")

            rows_this_page = 0
            page_results = []

            # Cerca tabella risultati
            result_table = None
            for table in soup.find_all('table'):
                headers = [th.get_text(strip=True).lower() for th in table.find_all('th')]
                if any(h in ' '.join(headers) for h in ['oggetto', 'numero', 'data', 'atto']):
                    result_table = table
                    break

            if result_table:
                header_cells = result_table.find('tr').find_all(['th', 'td']) if result_table.find('tr') else []
                col_headers = [c.get_text(strip=True).lower() for c in header_cells]
                print(f"[{name}] Colonne: {col_headers}")

                col_data = {}
                col_link = {}
                for i, h in enumerate(col_headers):
                    if 'data' in h and 'col_data' not in col_data:
                        col_data['data'] = i
                    if any(k in h for k in ['oggetto', 'titolo', 'descrizione']):
                        col_data['oggetto'] = i
                    if any(k in h for k in ['link', 'testo', 'visuali', 'allegat', 'dettaglio']):
                        col_data['link_col'] = i

                for row in result_table.find_all('tr')[1:]:
                    cells = row.find_all(['td', 'th'])
                    if not cells:
                        continue
                    rows_this_page += 1
                    cell_texts = [c.get_text(strip=True) for c in cells]
                    all_text = ' '.join(cell_texts)

                    date_val = cell_texts[col_data.get('data', 0)] if cell_texts else ''
                    oggetto_idx = col_data.get('oggetto', min(4, len(cells)-1))
                    oggetto = cell_texts[oggetto_idx] if oggetto_idx < len(cell_texts) else all_text

                    link_url = ''
                    link_col_idx = col_data.get('link_col', None)
                    if link_col_idx is not None and link_col_idx < len(cells):
                        a = cells[link_col_idx].find('a', href=True)
                        if a:
                            link_url = make_absolute_url(a['href'], base_url)
                    if not link_url:
                        for a in row.find_all('a', href=True):
                            href = a['href']
                            if href not in ('#', '', 'javascript:void(0)'):
                                link_url = make_absolute_url(href, base_url)
                                break

                    if matches_keywords(oggetto) or matches_keywords(all_text):
                        page_results.append({'data': date_val, 'oggetto': oggetto, 'link': link_url})

            print(f"[{name}] Pagina {page}: {rows_this_page} righe, {len(page_results)} match")
            results.extend(page_results)

            if rows_this_page == 0:
                break

            # Paginazione
            has_next = False
            for pat in [r'success', r'next', r'avanti', r'›', r'\bprossim']:
                nxt = soup.find('a', string=re.compile(pat, re.I))
                if nxt and nxt.get('href'):
                    has_next = True
                    break
            if not has_next:
                pag = soup.find(class_=re.compile(r'(pagination|paginat)', re.I))
                if pag:
                    all_pg = [a.get_text(strip=True) for a in pag.find_all('a') if a.get_text(strip=True).isdigit()]
                    if str(page + 1) in all_pg:
                        has_next = True
            if not has_next:
                break

            page += 1
            time.sleep(0.8)

    except Exception as e:
        print(f"[{name}] ERRORE: {e}")
        traceback.print_exc()
        PORTAL_ERRORS[name] = str(e)

    print(f"[{name}] Totale match: {len(results)}")
    return results


# ============================================================
# SCRAPER REGIONE CALABRIA
# ============================================================

def scrape_regione_calabria(date_from):
    name = 'Regione Calabria'
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    results = []
    print(f"\n[{name}] Data dal: {date_from}")

    try:
        page = 1
        while True:
            if page == 1:
                url = f"{base_url}?filter_date_from={date_from}"
            else:
                url = f"{base_url}?filter_date_from={date_from}&paged={page}"

            print(f"[{name}] Pagina {page}: {url}")
            resp = safe_get(url)
            if not resp:
                PORTAL_ERRORS[name] = "Portale non raggiungibile"
                break

            soup = BeautifulSoup(resp.text, 'lxml')
            if page == 1:
                print(f"[{name}] Titolo: {soup.title.string if soup.title else 'N/A'}")

            rows_this_page = 0
            page_results = []

            table = soup.find('table')
            if table:
                # Struttura nota:
                # [0] td: Tipologia
                # [1] td: Data Repertoriazione
                # [2] th: N (numero atto - usa th!)
                # [3] td: Dipartimento
                # [4] td: Oggetto (testo completo)
                # [5] td: Dettaglio (→ con link)
                rows = table.find_all('tr')
                header_row = rows[0] if rows else None
                if page == 1 and header_row:
                    headers = [c.get_text(strip=True) for c in header_row.find_all(['th', 'td'])]
                    print(f"[{name}] Colonne: {headers}")

                for row in rows[1:]:
                    cells = row.find_all(['td', 'th'])
                    if len(cells) < 3:
                        continue
                    rows_this_page += 1
                    cell_texts = [c.get_text(strip=True) for c in cells]

                    # Colonna data = indice 1
                    date_val = cell_texts[1] if len(cell_texts) > 1 else ''

                    # Colonna oggetto = indice 4 (include tutte le celle td+th)
                    oggetto = cell_texts[4] if len(cell_texts) > 4 else (cell_texts[-2] if len(cell_texts) >= 2 else '')

                    # Link nella cella dettaglio (ultima cella con link)
                    link_url = ''
                    for cell in reversed(cells):
                        a = cell.find('a', href=True)
                        if a and a.get('href', '') not in ('#', ''):
                            link_url = make_absolute_url(a['href'], base_url)
                            break

                    if matches_keywords(oggetto) or matches_keywords(' '.join(cell_texts)):
                        page_results.append({'data': date_val, 'oggetto': oggetto, 'link': link_url})

            print(f"[{name}] Pagina {page}: {rows_this_page} righe, {len(page_results)} match")
            results.extend(page_results)

            if rows_this_page == 0:
                break

            # Paginazione
            has_next = False
            next_a = soup.find('a', rel='next')
            if next_a:
                has_next = True
            if not has_next:
                pag = soup.find(class_=re.compile(r'(pagination|nav-links|page-numbers)', re.I))
                if pag:
                    nxt = pag.find('a', class_=re.compile(r'next', re.I))
                    if nxt:
                        has_next = True
                    else:
                        all_pg = [a.get_text(strip=True) for a in pag.find_all('a') if a.get_text(strip=True).isdigit()]
                        if str(page + 1) in all_pg:
                            has_next = True
            if not has_next:
                break

            page += 1
            time.sleep(0.8)

    except Exception as e:
        print(f"[{name}] ERRORE: {e}")
        traceback.print_exc()
        PORTAL_ERRORS[name] = str(e)

    print(f"[{name}] Totale match: {len(results)}")
    return results


# ============================================================
# SCRAPER TAR CATANZARO
# ============================================================

def scrape_tar_catanzaro(date_from):
    name = 'TAR Catanzaro'
    base_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
    results = []
    print(f"\n[{name}] Data dal: {date_from}")

    try:
        page = 1
        while True:
            if page == 1:
                url = f"{base_url}?publishDateFrom={date_from}"
            else:
                url = f"{base_url}?publishDateFrom={date_from}&page={page}"

            print(f"[{name}] Pagina {page}: {url}")
            resp = safe_get(url, verify=False)
            if not resp:
                PORTAL_ERRORS[name] = (
                    "Portale TAR non raggiungibile (503 / TLS error). "
                    "Il server giustizia-amministrativa.it restituisce errore SSL "
                    "dalla rete cloud. Eseguire lo script in locale."
                )
                break

            soup = BeautifulSoup(resp.text, 'lxml')
            if page == 1:
                print(f"[{name}] Titolo: {soup.title.string if soup.title else 'N/A'}")

            rows_this_page = 0
            page_results = []

            result_table = None
            for table in soup.find_all('table'):
                ths = [th.get_text(strip=True).lower() for th in table.find_all('th')]
                if any(k in ' '.join(ths) for k in ['parte', 'numero', 'tipo', 'data']):
                    result_table = table
                    break

            if result_table:
                header_cells = result_table.find('tr').find_all(['th', 'td']) if result_table.find('tr') else []
                col_headers = [c.get_text(strip=True).lower() for c in header_cells]
                print(f"[{name}] Colonne: {col_headers}")

                col_map = {}
                for i, h in enumerate(col_headers):
                    if 'data' in h and 'data' not in col_map:
                        col_map['data'] = i
                    if 'parte' in h or 'parti' in h:
                        col_map['parte'] = i
                    if any(k in h for k in ['tipo', 'provved', 'oggetto']):
                        col_map['tipo'] = i
                    if any(k in h for k in ['numero', 'num', 'n.']):
                        col_map['numero'] = i

                for row in result_table.find_all('tr')[1:]:
                    cells = row.find_all(['td', 'th'])
                    if not cells:
                        continue
                    rows_this_page += 1
                    cell_texts = [c.get_text(strip=True) for c in cells]
                    all_text = ' '.join(cell_texts)

                    date_val = cell_texts[col_map.get('data', 0)] if cell_texts else ''
                    parte = cell_texts[col_map.get('parte', 1)] if len(cell_texts) > col_map.get('parte', 1) else ''
                    tipo = cell_texts[col_map.get('tipo', 2)] if len(cell_texts) > col_map.get('tipo', 2) else ''
                    oggetto = f"{tipo} – {parte}".strip(' –') if (tipo or parte) else all_text

                    link_url = ''
                    for a in row.find_all('a', href=True):
                        href = a['href']
                        if href not in ('#', ''):
                            link_url = make_absolute_url(href, base_url)
                            break

                    if matches_keywords(parte) or matches_keywords(tipo) or matches_keywords(all_text):
                        page_results.append({'data': date_val, 'oggetto': oggetto, 'link': link_url})

            print(f"[{name}] Pagina {page}: {rows_this_page} righe, {len(page_results)} match")
            results.extend(page_results)

            if rows_this_page == 0:
                break

            has_next = False
            for pat in [r'success', r'next', r'avanti']:
                nxt = soup.find('a', string=re.compile(pat, re.I))
                if nxt and nxt.get('href'):
                    has_next = True
                    break
            if not has_next:
                pag = soup.find(class_=re.compile(r'(pagination|paginat)', re.I))
                if pag:
                    all_pg = [a.get_text(strip=True) for a in pag.find_all('a') if a.get_text(strip=True).isdigit()]
                    if str(page + 1) in all_pg:
                        has_next = True
            if not has_next:
                break

            page += 1
            time.sleep(0.8)

    except Exception as e:
        print(f"[{name}] ERRORE: {e}")
        traceback.print_exc()
        PORTAL_ERRORS[name] = str(e)

    print(f"[{name}] Totale match: {len(results)}")
    return results


# ============================================================
# CREAZIONE EXCEL
# ============================================================

HEADER_BG     = '2E75B6'
HEADER_FG     = 'FFFFFF'
ERROR_BG      = 'FFF2CC'
NODATA_BG     = 'F2F2F2'
LINK_COLOR    = '0563C1'
BORDER_COLOR  = 'BFBFBF'


def make_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def create_excel(sheets_data, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    hdr_font   = Font(name='Calibri', bold=True, size=11, color=HEADER_FG)
    hdr_fill   = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type='solid')
    hdr_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    link_font  = Font(name='Calibri', size=10, color=LINK_COLOR, underline='single')
    normal_font = Font(name='Calibri', size=10)
    wrap_align  = Alignment(vertical='top', wrap_text=True)
    err_fill    = PatternFill(start_color=ERROR_BG, end_color=ERROR_BG, fill_type='solid')

    sheet_order = [
        'ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
        'ASP Reggio Calabria', 'ASP Vibo Valentia',
        'Regione Calabria', 'TAR Catanzaro'
    ]

    for sheet_name in sheet_order:
        records = sheets_data.get(sheet_name, [])
        ws = wb.create_sheet(title=sheet_name[:31])

        # Header row
        col_labels = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
        for col, lbl in enumerate(col_labels, 1):
            cell = ws.cell(row=1, column=col, value=lbl)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = hdr_align
            cell.border = make_border()
        ws.row_dimensions[1].height = 22

        error_msg = PORTAL_ERRORS.get(sheet_name, '')

        if error_msg:
            # Portale non raggiungibile
            err_cell = ws.cell(row=2, column=1,
                               value=f"⚠ Portale non raggiungibile: {error_msg}")
            err_cell.font = Font(name='Calibri', size=10, color='7F0000', italic=True)
            err_cell.fill = err_fill
            ws.merge_cells('A2:D2')
            ws.row_dimensions[2].height = 40
            ws.cell(row=2, column=1).alignment = Alignment(wrap_text=True, vertical='top')

        elif not records:
            cell = ws.cell(row=2, column=1,
                           value='Nessun risultato per il periodo e le keyword indicate')
            cell.font = Font(name='Calibri', size=10, color='595959', italic=True)
            ws.merge_cells('A2:D2')

        else:
            for row_idx, record in enumerate(records, 2):
                data_val  = record.get('data', '')
                oggetto   = record.get('oggetto', '')
                descrizione = oggetto[:150] if oggetto else ''
                link_url  = record.get('link', '')

                c1 = ws.cell(row=row_idx, column=1, value=data_val)
                c1.font = normal_font; c1.alignment = wrap_align; c1.border = make_border()

                c2 = ws.cell(row=row_idx, column=2, value=oggetto)
                c2.font = normal_font; c2.alignment = wrap_align; c2.border = make_border()

                c3 = ws.cell(row=row_idx, column=3, value=descrizione)
                c3.font = normal_font; c3.alignment = wrap_align; c3.border = make_border()

                if link_url:
                    c4 = ws.cell(row=row_idx, column=4, value='Apri atto')
                    c4.hyperlink = link_url
                    c4.font = link_font
                else:
                    c4 = ws.cell(row=row_idx, column=4, value='N/D')
                    c4.font = normal_font
                c4.alignment = wrap_align
                c4.border = make_border()

                ws.row_dimensions[row_idx].height = max(15, min(80, len(oggetto) // 6))

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 65
        ws.column_dimensions['C'].width = 42
        ws.column_dimensions['D'].width = 12

        ws.freeze_panes = 'A2'
        last_row = max(2, len(records) + 1)
        ws.auto_filter.ref = f"A1:D{last_row}"

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")


# ============================================================
# MAIN
# ============================================================

def main():
    sheets_data = {}

    # --- ASP Portals ---
    for asp_name, asp_url in ASP_PORTALS.items():
        records = scrape_asp_portal(asp_name, asp_url, DATE_2DAYS_AGO)
        sheets_data[asp_name] = records
        time.sleep(1)

    # --- Regione Calabria ---
    sheets_data['Regione Calabria'] = scrape_regione_calabria(DATE_2DAYS_AGO)
    time.sleep(1)

    # --- TAR Catanzaro ---
    sheets_data['TAR Catanzaro'] = scrape_tar_catanzaro(DATE_3DAYS_AGO)

    # --- Riepilogo ---
    print("\n" + "=" * 60)
    print("RIEPILOGO")
    print("=" * 60)
    total = 0
    for name in ['ASP Cosenza','ASP Catanzaro','ASP Crotone',
                 'ASP Reggio Calabria','ASP Vibo Valentia',
                 'Regione Calabria','TAR Catanzaro']:
        recs = sheets_data.get(name, [])
        count = len(recs)
        total += count
        err = PORTAL_ERRORS.get(name, '')
        if err:
            status = f"⚠ Non raggiungibile"
        elif count:
            status = f"{count} atti trovati"
        else:
            status = "Nessun risultato"
        print(f"  {name:<25} {status}")
    print(f"  {'TOTALE MATCH':<25} {total}")
    print("=" * 60)

    create_excel(sheets_data, OUTPUT_FILE)
    print("\nCompletato.")
    return OUTPUT_FILE


if __name__ == '__main__':
    main()
