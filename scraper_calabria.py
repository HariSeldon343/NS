#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraping di: 5 portali ASP, Regione Calabria, TAR Catanzaro
Output: file XLSX con fogli per ogni portale
"""

import requests
import warnings
import re
import time
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlencode, urlparse, parse_qs
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# CONFIGURAZIONE
# ─────────────────────────────────────────────

TODAY = datetime.now()
DATE_FROM_2 = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3 = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
DATE_TO = TODAY.strftime('%Y-%m-%d')

KEYWORDS = [
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
# LIFE con word boundary separato
KEYWORD_LIFE = "LIFE"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.5',
    'Connection': 'keep-alive',
}

# ─────────────────────────────────────────────
# FILTRO KEYWORD
# ─────────────────────────────────────────────

def matches_keywords(text: str) -> bool:
    """Verifica se il testo contiene almeno una keyword (case-insensitive)."""
    if not text:
        return False
    text_up = text.upper()
    # LIFE con word boundary
    if re.search(r'\bLIFE\b', text, re.IGNORECASE):
        return True
    # tutte le altre keyword
    for kw in KEYWORDS:
        if kw.upper() in text_up:
            return True
    return False


# ─────────────────────────────────────────────
# REGIONE CALABRIA
# ─────────────────────────────────────────────

def scrape_regione_calabria():
    """Scarica tutti i provvedimenti della Regione Calabria negli ultimi 2 giorni."""
    print("\n=== Regione Calabria ===")
    results = []
    base = 'https://www.regione.calabria.it'
    session = requests.Session()
    session.headers.update(HEADERS)

    page = 1
    while True:
        if page == 1:
            url = f'{base}/provvedimenti-della-regione?filter_date_from={DATE_FROM_2}&filter_date_to={DATE_TO}&searchButton=Cerca&filter_active=true'
        else:
            url = f'{base}/provvedimenti-della-regione/?paged={page}&filter_date_from={DATE_FROM_2}&filter_date_to={DATE_TO}&searchButton=Cerca&filter_active=true'

        try:
            r = session.get(url, timeout=30)
            r.raise_for_status()
        except Exception as e:
            print(f"  Pagina {page}: errore → {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table', class_='table')
        if not table:
            print(f"  Pagina {page}: nessuna tabella trovata")
            break

        rows = table.find_all('tr')[1:]  # skip header
        if not rows:
            break

        found_on_page = 0
        for row in rows:
            tds = row.find_all('td')
            # Header ha 6 colonne (Tipologia, Data, N, Dipartimento, Oggetto, Dettaglio)
            # ma le righe dati ne hanno 5 (la colonna N è assente nei dati filtrati)
            if len(tds) < 5:
                continue
            data = tds[1].get_text(strip=True)
            # Se 5 col: [Tipologia, Data, Dipartimento, Oggetto, Link]
            # Se 6 col: [Tipologia, Data, N, Dipartimento, Oggetto, Link]
            if len(tds) >= 6:
                oggetto = tds[4].get_text(strip=True)
                link_tag = tds[5].find('a')
            else:
                oggetto = tds[3].get_text(strip=True)
                link_tag = tds[4].find('a')
            link = urljoin(base, link_tag['href']) if link_tag and link_tag.get('href') else ''

            if matches_keywords(oggetto):
                results.append({
                    'data': data,
                    'oggetto': oggetto,
                    'link': link,
                })
                found_on_page += 1

        print(f"  Pagina {page}: {len(rows)} atti, {found_on_page} match keyword")

        # Controlla se c'è pagina successiva
        pagination = soup.find('ul', class_='page-numbers')
        if not pagination:
            break
        next_link = pagination.find('a', string=re.compile('successiva|next|›|»', re.I))
        if not next_link:
            # Controlla se page+1 esiste
            all_page_links = pagination.find_all('a')
            page_nums = []
            for a in all_page_links:
                m = re.search(r'paged=(\d+)', a.get('href', ''))
                if m:
                    page_nums.append(int(m.group(1)))
            if page_nums and page < max(page_nums):
                page += 1
            else:
                break
        else:
            page += 1
        time.sleep(0.5)

    print(f"  Totale match: {len(results)}")
    return results


# ─────────────────────────────────────────────
# PORTALI ASP (5 portali)
# ─────────────────────────────────────────────

ASP_PORTALS = {
    'ASP Cosenza':        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria':'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}

def probe_asp_portal(name: str, url: str):
    """Tenta di accedere al portale ASP. Restituisce lista di risultati o None se irraggiungibile."""
    print(f"\n=== {name} ===")
    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        r = session.get(url, timeout=20, verify=False)
    except Exception as e:
        print(f"  ERRORE connessione: {e}")
        return None

    if r.status_code != 200:
        print(f"  HTTP {r.status_code} — portale non disponibile")
        return None

    soup = BeautifulSoup(r.text, 'lxml')
    return scrape_asp_portal(name, url, soup, session)


def scrape_asp_portal(name: str, base_url: str, soup_home: BeautifulSoup, session: requests.Session):
    """Scarica gli atti dal portale ASP iterando le pagine."""
    results = []

    # Individua form
    form = soup_home.find('form')
    if not form:
        print(f"  Nessun form trovato")
        return results

    action = form.get('action', base_url)
    if not action.startswith('http'):
        action = urljoin(base_url, action)

    method = (form.get('method') or 'get').lower()

    # Raccoglie tutti i campi
    fields = {}
    for inp in form.find_all(['input', 'select', 'textarea']):
        n = inp.get('name')
        if n:
            fields[n] = inp.get('value', '')

    # Imposta data pubblicazione
    date_field_id = 'dataPubblicazioneDal'
    for inp in form.find_all(['input', 'select']):
        if inp.get('id') == date_field_id or inp.get('name','').lower() in ('datapubblicazionedal', 'datadal', 'data_dal'):
            fields[inp.get('name', date_field_id)] = DATE_FROM_2
            break
    else:
        # Prova con il nome standard richiesto dall'utente
        fields['dataPubblicazioneDal'] = DATE_FROM_2

    page = 1
    while True:
        try:
            if method == 'post':
                r = session.post(action, data=fields, timeout=30, verify=False)
            else:
                r = session.get(action, params=fields, timeout=30, verify=False)
        except Exception as e:
            print(f"  Pagina {page}: errore → {e}")
            break

        if r.status_code != 200:
            print(f"  Pagina {page}: HTTP {r.status_code}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table')
        if not table:
            print(f"  Pagina {page}: nessuna tabella")
            break

        rows = table.find_all('tr')[1:]
        if not rows:
            break

        found_on_page = 0
        for row in rows:
            tds = row.find_all('td')
            if not tds:
                continue
            oggetto = ''
            link = ''
            data = ''
            # Prova a estrarre le colonne base
            for i, td in enumerate(tds):
                td_text = td.get_text(strip=True)
                td_a = td.find('a')
                # Euristica colonne
                if re.match(r'\d{2}/\d{2}/\d{4}', td_text) and not data:
                    data = td_text
                elif len(td_text) > 20 and not oggetto:
                    oggetto = td_text
                if td_a and td_a.get('href') and not link:
                    link = urljoin(base_url, td_a['href'])

            if not oggetto:
                oggetto = ' | '.join(td.get_text(strip=True) for td in tds)

            if matches_keywords(oggetto):
                results.append({'data': data, 'oggetto': oggetto, 'link': link})
                found_on_page += 1

        print(f"  Pagina {page}: {len(rows)} atti, {found_on_page} match keyword")

        # Paginazione
        next_page = soup.find('a', string=re.compile(r'successiv|next|›|»|\>', re.I))
        if not next_page:
            next_page = soup.find('a', title=re.compile(r'successiv|next', re.I))
        if next_page and next_page.get('href'):
            next_url = urljoin(base_url, next_page['href'])
            if method == 'post':
                # Probabilmente il link è GET con parametri
                r_next = session.get(next_url, timeout=30, verify=False)
                soup = BeautifulSoup(r_next.text, 'lxml')
                page += 1
                continue
            else:
                action = next_url
                fields = {}
                page += 1
        else:
            break

        time.sleep(0.3)

    print(f"  Totale match: {len(results)}")
    return results


# ─────────────────────────────────────────────
# TAR CATANZARO (Playwright)
# ─────────────────────────────────────────────

def scrape_tar_catanzaro():
    """
    Scarica i provvedimenti TAR Catanzaro usando Playwright.
    Nota: il backend Liferay del portale TAR blocca le richieste da IP non italiani
    (server in US/cloud). Il portale pubblico carica ma la ricerca restituisce errore.
    """
    print("\n=== TAR Catanzaro (Playwright) ===")
    results = []

    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        print("  Playwright non disponibile")
        return None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-dev-shm-usage'])
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            ignore_https_errors=True,
            locale='it-IT',
        )
        page = context.new_page()
        page.set_default_timeout(30000)

        try:
            print("  Caricamento pagina TAR...")
            page.goto('https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro', wait_until='networkidle')

            prefix_selector = '[id$="_publishDateFrom"]'
            page.wait_for_selector(prefix_selector, timeout=10000)
            page.fill(prefix_selector, DATE_FROM_3)
            print(f"  Data impostata: {DATE_FROM_3}")

            cerca = page.query_selector('button[id$="_search"]')
            if cerca:
                cerca.click()
            else:
                page.click('button:has-text("Cerca")')

            try:
                page.wait_for_load_state('networkidle', timeout=15000)
            except PWTimeout:
                pass
            time.sleep(2)

        except PWTimeout as e:
            print(f"  Timeout durante la navigazione: {e}")
            browser.close()
            return None

        html = page.content()
        soup = BeautifulSoup(html, 'lxml')

        # Controlla errore backend
        portlet = soup.find('div', class_=re.compile('portlet-boundary.*Jurisdictional'))
        if portlet:
            ptext = portlet.get_text()
            if 'errore' in ptext.lower() and 'Si è verificato' in ptext:
                print("  ERRORE BACKEND: il servizio di ricerca TAR restituisce errore.")
                print("  Causa probabile: IP non italiano (34.66.231.65, Google Cloud Iowa)")
                print("  Il portale TAR blocca le ricerche da IP cloud/extra-EU")
                browser.close()
                return None

        # Leggi risultati
        page_num = 1
        while True:
            html = page.content()
            soup = BeautifulSoup(html, 'lxml')
            portlet = soup.find('div', class_=re.compile('portlet-boundary.*Jurisdictional'))
            if not portlet:
                break

            tbody = portlet.find('tbody')
            if not tbody:
                break

            rows = tbody.find_all('tr')
            if not rows:
                break

            found_on_page = 0
            for row in rows:
                tds = row.find_all('td')
                if len(tds) < 4:
                    continue
                # Colonne TAR: NRG | Sezione | Parte | Tipo Udienza | Data Udienza | N. Provv | Data Pubb | Tipo | Relatore | Presidente | Esito
                # Indici approssimati: 0=NRG, 1=Sez, 2=Parte, 3=Data Pubb, ...
                nrg = tds[0].get_text(strip=True) if len(tds) > 0 else ''
                sezione = tds[1].get_text(strip=True) if len(tds) > 1 else ''
                parte = tds[2].get_text(strip=True) if len(tds) > 2 else ''
                data_pubb = ''
                link = ''
                for td in tds:
                    txt = td.get_text(strip=True)
                    if re.match(r'\d{2}/\d{2}/\d{4}', txt) and not data_pubb:
                        data_pubb = txt
                    a_tag = td.find('a')
                    if a_tag and a_tag.get('href') and not link:
                        href = a_tag['href']
                        if not href.startswith('http'):
                            href = 'https://www.giustizia-amministrativa.it' + href
                        link = href

                oggetto = parte  # filtro sulla colonna Parte per keyword
                if matches_keywords(oggetto):
                    results.append({'data': data_pubb, 'oggetto': f'{nrg} | {sezione} | {parte}', 'link': link})
                    found_on_page += 1

            print(f"  Pagina {page_num}: {len(rows)} atti, {found_on_page} match keyword")

            # Paginazione
            try:
                next_btn = page.query_selector('a:has-text("Successivo"), li.next a, .pagination .next a')
                if next_btn:
                    next_btn.click()
                    page.wait_for_load_state('networkidle')
                    time.sleep(1)
                    page_num += 1
                else:
                    break
            except Exception:
                break

        browser.close()

    print(f"  Totale match: {len(results)}")
    return results


# ─────────────────────────────────────────────
# CREAZIONE XLSX
# ─────────────────────────────────────────────

LINK_FONT = Font(color='0563C1', underline='single')

SHEET_HEADER_FILL = PatternFill('solid', fgColor='4472C4')
SHEET_HEADER_FONT = Font(bold=True, color='FFFFFF')

COL_WIDTHS = [15, 80, 55, 20]  # Data, Oggetto, Descrizione breve, Link

def write_sheet(ws, rows: list, sheet_name: str):
    """Scrive i dati in un foglio Excel."""
    headers = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']

    # Intestazione
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = SHEET_HEADER_FONT
        cell.fill = SHEET_HEADER_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS[col_idx - 1]

    ws.row_dimensions[1].height = 22

    if not rows:
        ws.cell(row=2, column=1, value='Nessun risultato')
        return

    for row_idx, item in enumerate(rows, 2):
        oggetto = item.get('oggetto', '')
        data = item.get('data', '')
        link = item.get('link', '')
        descr = oggetto[:150]

        ws.cell(row=row_idx, column=1, value=data)
        ws.cell(row=row_idx, column=2, value=oggetto).alignment = Alignment(wrap_text=True)
        ws.cell(row=row_idx, column=3, value=descr).alignment = Alignment(wrap_text=True)

        # Link come hyperlink cliccabile
        link_cell = ws.cell(row=row_idx, column=4, value='Apri' if link else '')
        if link:
            link_cell.hyperlink = link
            link_cell.font = LINK_FONT
            link_cell.alignment = Alignment(horizontal='center')

    ws.row_dimensions[1].height = 20


def build_excel(data_by_sheet: dict, output_path: str):
    """Crea il file Excel con un foglio per portale."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuovi foglio default

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
        rows = data_by_sheet.get(sheet_name, [])
        write_sheet(ws, rows, sheet_name)

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    data_by_sheet = {}

    IP_NOTA = (
        "PORTALE NON ACCESSIBILE DA IP ESTERO\n"
        "Ambiente di esecuzione: Google Cloud Iowa USA (IP 34.66.231.65)\n"
        "I portali ASP e TAR Calabria bloccano le ricerche da IP non italiani.\n"
        "Eseguire lo script localmente dalla rete italiana per ottenere i risultati."
    )

    # 1. Portali ASP
    for name, url in ASP_PORTALS.items():
        rows = probe_asp_portal(name, url)
        if rows is None:
            data_by_sheet[name] = [{'data': TODAY.strftime('%d/%m/%Y'),
                                     'oggetto': f'{IP_NOTA}\nURL portale: {url}',
                                     'link': url}]
        else:
            data_by_sheet[name] = rows

    # 2. Regione Calabria
    data_by_sheet['Regione Calabria'] = scrape_regione_calabria()

    # 3. TAR Catanzaro
    tar_results = scrape_tar_catanzaro()
    if tar_results is None:
        data_by_sheet['TAR Catanzaro'] = [{'data': TODAY.strftime('%d/%m/%Y'),
                                            'oggetto': f'{IP_NOTA}\nURL portale: https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro',
                                            'link': 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'}]
    else:
        data_by_sheet['TAR Catanzaro'] = tar_results

    # 4. Output Excel
    output_filename = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    output_path = f"/home/user/NS/{output_filename}"
    build_excel(data_by_sheet, output_path)

    # Riepilogo
    print("\n" + "="*60)
    print("RIEPILOGO")
    print("="*60)
    for sheet, rows in data_by_sheet.items():
        is_unavail = any('NON ACCESSIBILE' in r.get('oggetto', '') for r in rows)
        if is_unavail:
            print(f"  {sheet:<30} ⚠ NON ACCESSIBILE (IP estero)")
        else:
            print(f"  {sheet:<30} {len(rows)} match")
    print(f"\nFile: {output_path}")


if __name__ == '__main__':
    main()
