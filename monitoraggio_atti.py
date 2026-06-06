#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Raccoglie atti da portali ASP Calabria, Regione Calabria e TAR Catanzaro
e genera un file Excel con i risultati filtrati per keyword.

NOTA AMBIENTE CLOUD: I portali SISR (ASP) e TAR usano IP bloccati dall'egress
proxy Anthropic. Eseguire localmente su rete italiana per accesso completo.
"""

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import time
import sys
import urllib3
from urllib.parse import urljoin, urlparse

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# CONFIGURAZIONE DATE
# ============================================================
TODAY = datetime.now()
DATE_FROM_2 = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3 = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
DATE_TO = TODAY.strftime('%Y-%m-%d')

DATE_FROM_2_IT = (TODAY - timedelta(days=2)).strftime('%d/%m/%Y')
DATE_FROM_3_IT = (TODAY - timedelta(days=3)).strftime('%d/%m/%Y')
DATE_TO_IT = TODAY.strftime('%d/%m/%Y')

# ============================================================
# KEYWORD MATCHING
# ============================================================
KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]

LIFE_PATTERN = re.compile(r'\bLife\b', re.IGNORECASE)
KEYWORD_PATTERNS = [re.compile(re.escape(kw), re.IGNORECASE) for kw in KEYWORDS]


def matches_keywords(text):
    if not text:
        return False
    if LIFE_PATTERN.search(text):
        return True
    for pattern in KEYWORD_PATTERNS:
        if pattern.search(text):
            return True
    return False


# ============================================================
# HTTP SESSION
# ============================================================
HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/125.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}
TIMEOUT = 45

ERR_BLOCKED = "BLOCKED"
ERR_ERROR   = "ERROR"


def new_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.verify = False
    return s


# ============================================================
# PORTALI ASP SISR
# ============================================================

ASP_PORTALS = [
    {'name': 'ASP Cosenza',         'url': 'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'},
    {'name': 'ASP Catanzaro',       'url': 'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'},
    {'name': 'ASP Crotone',         'url': 'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'},
    {'name': 'ASP Reggio Calabria', 'url': 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'},
    {'name': 'ASP Vibo Valentia',   'url': 'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'},
]


def scrape_asp_portal(portal):
    """
    Scrape un portale ASP SISR Calabria (albo pretorio online).
    Parametri form: dataPubblicazioneDal / dataPubblicazioneAl (YYYY-MM-DD)
    Paginazione: parametro 'pagina' (1-based).
    """
    results = []
    error_type = None
    session = new_session()
    base_url = portal['url']
    name = portal['name']
    print(f"\n{'='*55}")
    print(f"[{name}] → {base_url}")

    # Passo 1: GET iniziale per hidden fields
    form_action = base_url
    form_method = 'post'
    hidden_fields = {}

    try:
        resp = session.get(base_url, timeout=TIMEOUT)
        if resp.status_code == 503:
            print(f"  503 Service Unavailable - portale non raggiungibile")
            error_type = ERR_BLOCKED
            return [], error_type
        resp.raise_for_status()
        soup0 = BeautifulSoup(resp.text, 'lxml')
        form = soup0.find('form')
        if form:
            if form.get('action'):
                form_action = urljoin(base_url, form['action'])
            form_method = form.get('method', 'post').lower()
            for h in form.find_all('input', type='hidden'):
                if h.get('name'):
                    hidden_fields[h['name']] = h.get('value', '')
        print(f"  Form: {form_method.upper()} {form_action}")
    except requests.exceptions.HTTPError as e:
        if '503' in str(e):
            print(f"  503 - portale non raggiungibile")
            error_type = ERR_BLOCKED
            return [], error_type
        print(f"  WARN GET iniziale: {e}")
    except Exception as e:
        print(f"  WARN GET iniziale: {e}")
        error_type = ERR_ERROR

    # Passo 2: iterazione pagine
    page_num = 1
    while True:
        payload = {
            **hidden_fields,
            'dataPubblicazioneDal': DATE_FROM_2,
            'dataPubblicazioneAl':  DATE_TO,
            'pagina': str(page_num),
        }
        print(f"  → Pagina {page_num}...", end=' ', flush=True)
        try:
            if form_method == 'post':
                resp = session.post(form_action, data=payload, timeout=TIMEOUT)
            else:
                resp = session.get(form_action, params=payload, timeout=TIMEOUT)
            if resp.status_code == 503:
                print("503 - bloccato")
                if error_type is None:
                    error_type = ERR_BLOCKED
                break
            resp.raise_for_status()
        except Exception as e:
            print(f"ERRORE: {e}")
            if error_type is None:
                error_type = ERR_ERROR
            break

        soup = BeautifulSoup(resp.text, 'lxml')
        page_rows = _parse_asp_table(soup, base_url)
        print(f"{len(page_rows)} righe")

        if not page_rows:
            break
        results.extend(page_rows)
        if not _has_next(soup, page_num):
            break
        page_num += 1
        time.sleep(0.7)

    print(f"  Totale grezzo: {len(results)}")
    filtered = [r for r in results if matches_keywords(r.get('oggetto', ''))]
    print(f"  Matchati:      {len(filtered)}")
    return filtered, error_type


def _parse_asp_table(soup, base_url):
    """Estrae righe dalla tabella risultati albo SISR."""
    rows_out = []
    for table in soup.find_all('table'):
        trs = table.find_all('tr')
        if len(trs) < 2:
            continue
        hdrs = [c.get_text(strip=True).lower() for c in trs[0].find_all(['th', 'td'])]
        date_idx    = next((i for i, h in enumerate(hdrs) if 'data' in h and 'pub' in h), None)
        if date_idx is None:
            date_idx = next((i for i, h in enumerate(hdrs) if 'data' in h), None)
        obj_idx = next((i for i, h in enumerate(hdrs)
                        if any(k in h for k in ('oggetto', 'descr', 'titolo', 'atto'))), None)
        if obj_idx is None and hdrs:
            obj_idx = min(1, len(hdrs) - 1)

        for tr in trs[1:]:
            cells = tr.find_all(['td', 'th'])
            if not cells:
                continue
            data    = cells[date_idx].get_text(strip=True) if date_idx is not None and date_idx < len(cells) else ''
            oggetto = cells[obj_idx].get_text(strip=True)  if obj_idx  is not None and obj_idx  < len(cells) else ''
            link_url = base_url
            for a in tr.find_all('a', href=True):
                href = a['href']
                if href and href not in ('#',):
                    link_url = urljoin(base_url, href)
                    break
            if oggetto:
                rows_out.append({'data': data, 'oggetto': oggetto, 'link': link_url})
    return rows_out


def _has_next(soup, current_page):
    for a in soup.find_all('a', href=True):
        text = a.get_text(strip=True).lower()
        href = a['href']
        if text in ('>', '>>', 'successiva', 'next', 'avanti', '→', 'successivo'):
            return True
        for p in (f'pagina={current_page+1}', f'page={current_page+1}', f'p={current_page+1}'):
            if p in href:
                return True
    return False


# ============================================================
# REGIONE CALABRIA
# ============================================================

def scrape_regione_calabria():
    """
    Scrape https://www.regione.calabria.it/provvedimenti-della-regione/
    Tabella: Tipologia | Data Repertoriazione | N | Dipartimento | Oggetto | Dettaglio(link)
    Paginazione WordPress: ?paged=N  (8 righe/pagina)
    """
    results = []
    error_type = None
    session = new_session()
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    print(f"\n{'='*55}")
    print(f"[Regione Calabria] → {base_url}")

    page_num = 1
    while True:
        params = {
            'filter_date_from': DATE_FROM_2,
            'filter_date_to':   DATE_TO,
            'paged':            str(page_num),
        }
        print(f"  → Pagina {page_num}...", end=' ', flush=True)
        try:
            resp = session.get(base_url, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
        except requests.exceptions.HTTPError as e:
            print(f"ERRORE HTTP: {e}")
            error_type = ERR_ERROR
            break
        except Exception as e:
            print(f"ERRORE: {e}")
            error_type = ERR_ERROR
            break

        soup = BeautifulSoup(resp.text, 'lxml')
        page_rows = _parse_regione_table(soup)
        print(f"{len(page_rows)} righe")

        if not page_rows:
            break
        results.extend(page_rows)
        if not _has_next(soup, page_num):
            break
        page_num += 1
        time.sleep(0.7)

    print(f"  Totale grezzo: {len(results)}")
    filtered = [r for r in results if matches_keywords(r.get('oggetto', ''))]
    print(f"  Matchati:      {len(filtered)}")
    return filtered, error_type


def _parse_regione_table(soup):
    """
    Estrae atti dalla tabella Regione Calabria.
    Colonne attese: Tipologia | Data Repertoriazione | N | Dipartimento | Oggetto | Dettaglio
    """
    rows_out = []
    for table in soup.find_all('table'):
        trs = table.find_all('tr')
        if len(trs) < 2:
            continue
        hdrs = [c.get_text(strip=True).lower() for c in trs[0].find_all(['th', 'td'])]
        if not any('oggetto' in h or 'provvediment' in h for h in hdrs):
            continue

        # Individua colonne per nome
        date_idx   = next((i for i, h in enumerate(hdrs) if 'data' in h), 1)
        obj_idx    = next((i for i, h in enumerate(hdrs) if 'oggetto' in h or 'provvediment' in h), 4)
        detail_idx = next((i for i, h in enumerate(hdrs) if 'dettaglio' in h or 'link' in h or '→' in hdrs[i]), len(hdrs) - 1)

        for tr in trs[1:]:
            cells = tr.find_all(['td', 'th'])
            if not cells or len(cells) < max(date_idx, obj_idx) + 1:
                continue
            data    = cells[date_idx].get_text(strip=True) if date_idx < len(cells) else ''
            oggetto = cells[obj_idx].get_text(strip=True)  if obj_idx  < len(cells) else ''

            # Link dal "Dettaglio" (→)
            link_url = ''
            for c in cells:
                for a in c.find_all('a', href=True):
                    href = a['href']
                    if href.startswith('http'):
                        link_url = href
                        break
                if link_url:
                    break

            if oggetto:
                rows_out.append({'data': data, 'oggetto': oggetto, 'link': link_url or 'https://www.regione.calabria.it/provvedimenti-della-regione/'})

    return rows_out


# ============================================================
# TAR CATANZARO
# ============================================================

def scrape_tar_catanzaro():
    """
    Scrape https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro
    Filtro: publishDateFrom (ultimi 3 giorni)
    Filtra su colonna "Parte" con le stesse keyword.
    """
    results = []
    error_type = None
    session = new_session()
    base_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
    print(f"\n{'='*55}")
    print(f"[TAR Catanzaro] → {base_url}")

    page_num = 1
    while True:
        params = {
            'publishDateFrom': DATE_FROM_3,
            'publishDateTo':   DATE_TO,
            'page':            str(page_num),
        }
        print(f"  → Pagina {page_num}...", end=' ', flush=True)
        try:
            resp = session.get(base_url, params=params, timeout=TIMEOUT)
            if resp.status_code == 503:
                print("503 - bloccato")
                error_type = ERR_BLOCKED
                break
            resp.raise_for_status()
        except requests.exceptions.HTTPError as e:
            if '503' in str(e):
                print("503 - bloccato")
                error_type = ERR_BLOCKED
            else:
                print(f"ERRORE: {e}")
                error_type = ERR_ERROR
            break
        except Exception as e:
            print(f"ERRORE: {e}")
            error_type = ERR_ERROR
            break

        soup = BeautifulSoup(resp.text, 'lxml')
        page_rows = _parse_tar_table(soup, base_url)
        print(f"{len(page_rows)} righe")

        if not page_rows:
            break
        results.extend(page_rows)
        if not _has_next(soup, page_num):
            break
        page_num += 1
        time.sleep(0.7)

    print(f"  Totale grezzo: {len(results)}")
    filtered = [r for r in results
                if matches_keywords(r.get('parte', '')) or matches_keywords(r.get('oggetto', ''))]
    print(f"  Matchati:      {len(filtered)}")
    return filtered, error_type


def _parse_tar_table(soup, base_url):
    rows_out = []
    for table in soup.find_all('table'):
        trs = table.find_all('tr')
        if len(trs) < 2:
            continue
        hdrs = [c.get_text(strip=True).lower() for c in trs[0].find_all(['th', 'td'])]
        date_idx  = next((i for i, h in enumerate(hdrs) if 'data' in h), None)
        parte_idx = next((i for i, h in enumerate(hdrs) if 'parte' in h), None)
        tipo_idx  = next((i for i, h in enumerate(hdrs)
                          if any(k in h for k in ('tipo', 'oggetto', 'provvedimento'))), None)

        for tr in trs[1:]:
            cells = tr.find_all(['td', 'th'])
            if not cells:
                continue
            data    = cells[date_idx].get_text(strip=True)  if date_idx  is not None and date_idx  < len(cells) else ''
            parte   = cells[parte_idx].get_text(strip=True) if parte_idx is not None and parte_idx < len(cells) else ''
            oggetto = cells[tipo_idx].get_text(strip=True)  if tipo_idx  is not None and tipo_idx  < len(cells) else ''
            link_url = base_url
            for a in tr.find_all('a', href=True):
                href = a['href']
                if href and href != '#':
                    link_url = urljoin('https://www.giustizia-amministrativa.it', href)
                    break
            testo = (parte + ' ' + oggetto).strip()
            if testo:
                rows_out.append({'data': data, 'oggetto': oggetto or parte, 'parte': parte, 'link': link_url})
    return rows_out


# ============================================================
# EXCEL OUTPUT
# ============================================================

SHEET_ORDER = [
    'ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
    'ASP Reggio Calabria', 'ASP Vibo Valentia',
    'Regione Calabria', 'TAR Catanzaro',
]

COL_HEADERS = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
COL_WIDTHS  = [14, 65, 42, 16]


def _apply_border(cell, color='AAAAAA'):
    side = Side(border_style='thin', color=color)
    cell.border = Border(left=side, right=side, top=side, bottom=side)


def _header_cell(cell, value):
    cell.value = value
    cell.font      = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
    cell.fill      = PatternFill(start_color='003399', end_color='003399', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    _apply_border(cell, '003366')


def _data_cell(cell, value, link=None, italic=False, color=None):
    cell.value     = value
    cell.alignment = Alignment(vertical='top', wrap_text=True)
    if link:
        cell.hyperlink = link
        cell.font = Font(color='0563C1', underline='single', size=10, name='Calibri')
    elif italic:
        cell.font = Font(italic=True, size=10, name='Calibri', color=color or '555555')
    else:
        cell.font = Font(size=10, name='Calibri')
    _apply_border(cell)


def create_excel(all_data, all_errors, output_path):
    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    for sheet_name in SHEET_ORDER:
        records   = all_data.get(sheet_name, [])
        err_type  = all_errors.get(sheet_name)
        ws = wb.create_sheet(title=sheet_name)

        # Intestazione
        for col_num, h in enumerate(COL_HEADERS, 1):
            _header_cell(ws.cell(row=1, column=col_num), h)
        ws.row_dimensions[1].height = 28

        row = 2
        if err_type == ERR_BLOCKED:
            msg = (
                "⚠ Portale non raggiungibile dall'ambiente cloud (HTTP 503). "
                "Eseguire lo script localmente su rete italiana per accesso completo."
            )
            _data_cell(ws.cell(row=row, column=1), msg, italic=True, color='CC0000')
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        elif not records:
            _data_cell(ws.cell(row=row, column=1), 'Nessun risultato nel periodo selezionato', italic=True)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        else:
            for rec in records:
                data    = rec.get('data', '')
                oggetto = rec.get('oggetto', '')
                descr   = oggetto[:150]
                link    = rec.get('link', '')
                _data_cell(ws.cell(row=row, column=1), data)
                _data_cell(ws.cell(row=row, column=2), oggetto)
                _data_cell(ws.cell(row=row, column=3), descr)
                _data_cell(ws.cell(row=row, column=4), 'Apri atto', link=link if link else None)
                if not link:
                    _data_cell(ws.cell(row=row, column=4), 'N/D', italic=True)
                row += 1

        for col_num, width in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = width
        ws.freeze_panes = 'A2'

    wb.save(output_path)
    print(f"\n✓ Excel salvato: {output_path}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 55)
    print(" MONITORAGGIO ATTI CALABRIA")
    print(f" Esecuzione: {TODAY.strftime('%d/%m/%Y %H:%M')}")
    print(f" Date ASP/Regione: {DATE_FROM_2} → {DATE_TO}")
    print(f" Date TAR:         {DATE_FROM_3} → {DATE_TO}")
    print("=" * 55)

    all_data   = {}
    all_errors = {}

    # --- ASP SISR ---
    for portal in ASP_PORTALS:
        name = portal['name']
        try:
            records, err = scrape_asp_portal(portal)
            all_data[name]   = records
            all_errors[name] = err
        except Exception as e:
            print(f"  ERRORE GRAVE [{name}]: {e}")
            all_data[name]   = []
            all_errors[name] = ERR_ERROR

    # --- Regione Calabria ---
    try:
        records, err = scrape_regione_calabria()
        all_data['Regione Calabria']   = records
        all_errors['Regione Calabria'] = err
    except Exception as e:
        print(f"  ERRORE GRAVE [Regione Calabria]: {e}")
        all_data['Regione Calabria']   = []
        all_errors['Regione Calabria'] = ERR_ERROR

    # --- TAR Catanzaro ---
    try:
        records, err = scrape_tar_catanzaro()
        all_data['TAR Catanzaro']   = records
        all_errors['TAR Catanzaro'] = err
    except Exception as e:
        print(f"  ERRORE GRAVE [TAR Catanzaro]: {e}")
        all_data['TAR Catanzaro']   = []
        all_errors['TAR Catanzaro'] = ERR_ERROR

    # --- Excel ---
    filename    = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    output_path = f"/home/user/NS/{filename}"
    create_excel(all_data, all_errors, output_path)

    # --- Riepilogo ---
    print("\n" + "=" * 55)
    print(" RIEPILOGO FINALE")
    print("=" * 55)
    for name in SHEET_ORDER:
        err = all_errors.get(name)
        n   = len(all_data.get(name, []))
        if err == ERR_BLOCKED:
            status = "⚠ Non raggiungibile (503 - eseguire localmente)"
        elif err == ERR_ERROR:
            status = "✗ Errore di connessione"
        elif n > 0:
            status = f"✓ {n} atti con match keyword"
        else:
            status = "✓ Nessun match nel periodo"
        print(f"  {name:<25} {status}")
    print("=" * 55)
    print(f"  File: {output_path}")


if __name__ == '__main__':
    main()
