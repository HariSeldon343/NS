#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoraggio Atti Calabria - Scraper multi-portale
Recupera atti da portali ASP Calabria, Regione Calabria e TAR Catanzaro.
Filtra per keyword e produce file Excel.
"""

import requests
import warnings
import time
import re
import logging
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIGURAZIONE
# ─────────────────────────────────────────────
DATE_FROM_ASP = "2026-06-03"        # ultimi 2 giorni per ASP
DATE_FROM_ASP_IT = "03/06/2026"     # formato italiano per ASP
DATE_FROM_REGCAL = "2026-06-03"     # ultimi 2 giorni per Regione Calabria
DATE_FROM_TAR = "2026-06-02"        # ultimi 3 giorni per TAR
DATE_FROM_TAR_IT = "02/06/2026"     # formato italiano per TAR

OUTPUT_FILE = "/home/user/NS/Monitoraggio_Atti_Calabria_05-06-2026.xlsx"

KEYWORDS = [
    r"\bADI\b",
    r"Assistenza domiciliare",
    r"\bANMIC\b",
    r"Accreditamento",
    r"Aumento di budget",
    r"Autismo",
    r"Autorizzazione all[''']esercizio",
    r"Autorizzazione alla realizzazione",
    r"Autorizzazioni",
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

# Compila le regex una sola volta
KEYWORDS_RE = [re.compile(kw, re.IGNORECASE) for kw in KEYWORDS]

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}

ASP_PORTALS = [
    {
        "name": "ASP Cosenza",
        "sheet": "ASP Cosenza",
        "url": "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
        "code": "CO",
    },
    {
        "name": "ASP Catanzaro",
        "sheet": "ASP Catanzaro",
        "url": "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
        "code": "CZ",
    },
    {
        "name": "ASP Crotone",
        "sheet": "ASP Crotone",
        "url": "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
        "code": "KR",
    },
    {
        "name": "ASP Reggio Calabria",
        "sheet": "ASP Reggio Calabria",
        "url": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
        "code": "RC",
    },
    {
        "name": "ASP Vibo Valentia",
        "sheet": "ASP Vibo Valentia",
        "url": "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
        "code": "VV",
    },
]


# ─────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────
def match_keywords(text: str) -> bool:
    """Restituisce True se il testo contiene almeno una keyword."""
    for pattern in KEYWORDS_RE:
        if pattern.search(text):
            return True
    return False


def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS_HTTP)
    session.verify = False
    return session


def safe_get(session: requests.Session, url: str, params=None, timeout=25):
    """GET con gestione errori; restituisce Response o None."""
    try:
        resp = session.get(url, params=params, timeout=timeout, allow_redirects=True)
        resp.raise_for_status()
        return resp
    except requests.exceptions.HTTPError as e:
        log.warning(f"HTTP {e.response.status_code} per {url}")
        return None
    except requests.exceptions.Timeout:
        log.warning(f"Timeout per {url}")
        return None
    except requests.exceptions.ConnectionError as e:
        log.warning(f"Connessione fallita per {url}: {e}")
        return None
    except Exception as e:
        log.warning(f"Errore generico per {url}: {e}")
        return None


# ─────────────────────────────────────────────
# SCRAPER: PORTALI ASP
# ─────────────────────────────────────────────
def scrape_asp_portal(portal: dict, session: requests.Session) -> list:
    """
    Scrapa un portale ASP SISR.
    Struttura: form GET/POST con parametro dataPubblicazioneDal.
    Tabella risultati con paginazione.
    Restituisce lista di dict: {data, oggetto, link}
    """
    name = portal["name"]
    base_url = portal["url"]
    log.info(f"Scraping {name} ...")

    results = []
    page = 1
    max_pages = 50  # sicurezza

    while page <= max_pages:
        # Prova prima GET con parametri
        params_get = {
            "dataPubblicazioneDal": DATE_FROM_ASP_IT,
            "page": str(page),
        }
        resp = safe_get(session, base_url, params=params_get, timeout=30)

        if resp is None:
            # Prova formato alternativo della data
            params_alt = {
                "dataPubblicazioneDal": DATE_FROM_ASP,
                "page": str(page),
            }
            resp = safe_get(session, base_url, params=params_alt, timeout=30)

        if resp is None:
            # Prova POST
            try:
                post_data = {
                    "dataPubblicazioneDal": DATE_FROM_ASP_IT,
                    "page": str(page),
                }
                resp_post = session.post(base_url, data=post_data, timeout=30)
                resp_post.raise_for_status()
                resp = resp_post
            except Exception as e:
                log.warning(f"{name} - Pagina {page}: tutti i tentativi falliti ({e})")
                break

        if resp is None or resp.status_code not in (200, 206):
            log.warning(f"{name} - Pagina {page}: HTTP {resp.status_code if resp else 'N/A'}")
            break

        soup = BeautifulSoup(resp.text, "html.parser")

        # Cerca la tabella dei risultati
        # Pattern tipico SISR: tabella con class "tabella" o "table" o id "albo"
        table = (
            soup.find("table", {"id": re.compile(r"albo|risultati|atti", re.I)})
            or soup.find("table", {"class": re.compile(r"albo|risultati|atti|table", re.I)})
            or soup.find("table")
        )

        if not table:
            log.warning(f"{name} - Pagina {page}: nessuna tabella trovata")
            break

        rows = table.find_all("tr")
        data_rows = [r for r in rows if r.find("td")]

        if not data_rows:
            log.info(f"{name} - Pagina {page}: nessun dato (fine paginazione)")
            break

        page_count = 0
        for row in data_rows:
            cells = row.find_all("td")
            if not cells:
                continue

            # Estrai oggetto e data - le colonne possono variare tra portali
            # Struttura tipica SISR: Numero | Data | Oggetto | Ente | Link
            all_text = " ".join(c.get_text(separator=" ", strip=True) for c in cells)
            oggetto = ""
            data_pubbl = ""
            link = ""

            # Cerca link nell'intera riga
            link_el = row.find("a", href=True)
            if link_el:
                href = link_el.get("href", "")
                if href and not href.startswith("#"):
                    if href.startswith("http"):
                        link = href
                    else:
                        from urllib.parse import urljoin
                        link = urljoin(base_url, href)

            # Cerca data (pattern gg/mm/aaaa o aaaa-mm-gg)
            date_pattern = re.search(
                r"\b(\d{2}[/\-]\d{2}[/\-]\d{4}|\d{4}[/\-]\d{2}[/\-]\d{2})\b", all_text
            )
            if date_pattern:
                data_pubbl = date_pattern.group(1)

            # Prendi la cella più lunga come oggetto (escludendo celle con solo date/numeri)
            for cell in cells:
                txt = cell.get_text(separator=" ", strip=True)
                if len(txt) > len(oggetto) and not re.fullmatch(r"[\d/\-\s,\.]+", txt):
                    oggetto = txt

            # Fallback: usa tutto il testo come oggetto
            if not oggetto:
                oggetto = all_text[:300]

            if oggetto and match_keywords(oggetto):
                results.append({
                    "data": data_pubbl,
                    "oggetto": oggetto,
                    "link": link,
                })
                page_count += 1

        log.info(f"{name} - Pagina {page}: {len(data_rows)} righe, {page_count} match")

        # Controlla paginazione
        next_page = None
        # Cerca link "successiva", "next", ">" o numero pagina
        pag_links = soup.find_all("a", href=True)
        for pl in pag_links:
            href = pl.get("href", "")
            txt = pl.get_text(strip=True).lower()
            if txt in ("successiva", "next", ">", "»", "avanti"):
                next_page = href
                break
            if re.search(rf"[?&]page={page+1}", href):
                next_page = href
                break

        if not next_page:
            # Cerca anche per pattern "pagina N" o numero
            pag_items = soup.find_all(
                class_=re.compile(r"paginat|page-num|current", re.I)
            )
            # Se non c'è nessun link "avanti", siamo all'ultima pagina
            has_next = any(
                "successiv" in p.get_text(strip=True).lower()
                or "next" in p.get_text(strip=True).lower()
                or ">" in p.get_text(strip=True)
                for p in pag_items
            )
            if not has_next and page > 1:
                break

        page += 1
        time.sleep(0.5)

    log.info(f"{name}: totale {len(results)} atti con keyword")
    return results


# ─────────────────────────────────────────────
# SCRAPER: REGIONE CALABRIA
# ─────────────────────────────────────────────
def scrape_regione_calabria(session: requests.Session) -> list:
    """
    Scrapa il portale Regione Calabria provvedimenti.
    URL: https://www.regione.calabria.it/provvedimenti-della-regione/
    Parametro: filter_date_from=2026-06-03
    Tabella con class 'table-striped'.
    Struttura righe dati (5 celle): [Tipologia][Data][Dipartimento][Oggetto][→link]
    Struttura righe dati (6 celle): [Tipologia][Data][N][Dipartimento][Oggetto][→link]
    Paginazione: ?paged=N&filter_date_from=...
    """
    log.info("Scraping Regione Calabria ...")
    base_url = "https://www.regione.calabria.it/provvedimenti-della-regione/"
    results = []
    page = 1
    max_pages = 100

    while page <= max_pages:
        if page == 1:
            params = {
                "filter_date_from": DATE_FROM_REGCAL,
                "searchButton": "Cerca",
            }
        else:
            params = {
                "paged": str(page),
                "filter_date_from": DATE_FROM_REGCAL,
            }

        resp = safe_get(session, base_url, params=params)
        if resp is None:
            log.warning(f"Regione Calabria - Pagina {page}: errore")
            break

        soup = BeautifulSoup(resp.text, "html.parser")

        # Tabella principale
        table = soup.find("table", {"class": re.compile(r"table", re.I)})
        if not table:
            log.warning(f"Regione Calabria - Pagina {page}: nessuna tabella")
            break

        rows = table.find_all("tr")
        data_rows = [r for r in rows if r.find("td")]

        if not data_rows:
            log.info(f"Regione Calabria - Pagina {page}: fine dati")
            break

        page_match = 0
        for row in data_rows:
            cells = row.find_all("td")
            n_cells = len(cells)

            if n_cells < 4:
                continue

            tipologia = cells[0].get_text(strip=True)
            data_pubbl = cells[1].get_text(strip=True)

            # Struttura flessibile: la colonna N può essere presente o assente
            # nelle righe dati (il portale rimuove la cella N nelle righe dati)
            if n_cells == 5:
                # [tipo][data][dipartimento][oggetto][→]
                dipartimento = cells[2].get_text(strip=True)
                oggetto = cells[3].get_text(strip=True)
                link_cell = cells[4]
            elif n_cells == 6:
                # [tipo][data][N][dipartimento][oggetto][→]
                dipartimento = cells[3].get_text(strip=True)
                oggetto = cells[4].get_text(strip=True)
                link_cell = cells[5]
            else:
                # Fallback: usa la penultima cella come oggetto
                dipartimento = cells[-3].get_text(strip=True) if n_cells >= 3 else ""
                oggetto = cells[-2].get_text(strip=True)
                link_cell = cells[-1]

            # Link nella cella dettaglio (→)
            link = ""
            link_el = link_cell.find("a", href=True)
            if link_el:
                link = link_el.get("href", "")

            # Match keyword su OGGETTO (campo primario) e DIPARTIMENTO
            search_text = f"{oggetto} {dipartimento}"

            if match_keywords(search_text):
                results.append({
                    "data": data_pubbl,
                    "oggetto": oggetto,
                    "link": link,
                })
                page_match += 1
                log.debug(f"  MATCH: {data_pubbl} | {oggetto[:80]}")

        log.info(f"Regione Calabria - Pagina {page}: {len(data_rows)} righe, {page_match} match")

        # Controlla se c'è pagina successiva
        next_links = soup.find_all("a", class_=re.compile(r"page-numbers", re.I))
        has_next = any(
            "successiv" in l.get_text(strip=True).lower()
            or "Pagina successiva" in l.get_text(strip=True)
            for l in next_links
        )

        if not has_next:
            log.info(f"Regione Calabria: ultima pagina = {page}")
            break

        page += 1
        time.sleep(0.6)

    log.info(f"Regione Calabria: totale {len(results)} atti con keyword")
    return results


# ─────────────────────────────────────────────
# SCRAPER: TAR CATANZARO
# ─────────────────────────────────────────────
def scrape_tar_catanzaro(session: requests.Session) -> list:
    """
    Scrapa il TAR di Catanzaro (giustizia-amministrativa.it).
    URL: https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro
    Parametro: publishDateFrom=02/06/2026
    Filtra sulla colonna 'Parte' per keyword.
    """
    log.info("Scraping TAR Catanzaro ...")
    base_url = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
    results = []

    # Pattern di parametri da provare
    param_variants = [
        {"publishDateFrom": DATE_FROM_TAR_IT},
        {"publishDateFrom": DATE_FROM_TAR},
        {"dataDal": DATE_FROM_TAR_IT},
        {"dataInizio": DATE_FROM_TAR_IT},
    ]

    resp = None
    for params in param_variants:
        resp = safe_get(session, base_url, params=params, timeout=30)
        if resp is not None:
            log.info(f"TAR: connessione riuscita con params={params}")
            break

    if resp is None:
        # Prova senza parametri (pagina base)
        resp = safe_get(session, base_url, timeout=30)

    if resp is None:
        log.warning("TAR Catanzaro: impossibile connettersi")
        return results

    page = 1
    max_pages = 50

    while page <= max_pages:
        if page > 1:
            # Richiesta pagine successive
            params_pag = {
                "publishDateFrom": DATE_FROM_TAR_IT,
                "p_r_p_page": str(page),  # Liferay pagination
            }
            resp = safe_get(session, base_url, params=params_pag, timeout=30)
            if resp is None:
                break

        soup = BeautifulSoup(resp.text, "html.parser")

        # Cerca tabella dei provvedimenti
        # Il portale giustizia-amministrativa usa tipicamente una tabella con "Parte"
        table = (
            soup.find("table", {"id": re.compile(r"provved|atti|result", re.I)})
            or soup.find("table", {"class": re.compile(r"result|provved|atti|table", re.I)})
            or soup.find("table")
        )

        if not table:
            # Cerca lista alternativa
            items = soup.find_all(
                class_=re.compile(r"provved|atto|sentenza|decreto|result-item", re.I)
            )
            if not items:
                log.warning(f"TAR - Pagina {page}: nessuna tabella o lista trovata")
                break

            # Estrai da lista
            for item in items:
                text = item.get_text(separator=" ", strip=True)
                link_el = item.find("a", href=True)
                link = link_el.get("href", "") if link_el else ""
                if not link.startswith("http"):
                    from urllib.parse import urljoin
                    link = urljoin(base_url, link)

                date_m = re.search(r"\b(\d{2}[/\-]\d{2}[/\-]\d{4})\b", text)
                data_pubbl = date_m.group(1) if date_m else ""

                if match_keywords(text):
                    results.append({
                        "data": data_pubbl,
                        "oggetto": text[:500],
                        "link": link,
                    })
        else:
            rows = table.find_all("tr")
            data_rows = [r for r in rows if r.find("td")]

            if not data_rows:
                break

            # Identifica indice colonne dalla riga header
            header_row = table.find("tr", {"class": re.compile(r"header|head", re.I)}) or rows[0]
            header_cells = header_row.find_all(["th", "td"])
            col_map = {}
            for i, hcell in enumerate(header_cells):
                htxt = hcell.get_text(strip=True).lower()
                if "data" in htxt or "pubblic" in htxt:
                    col_map["data"] = i
                elif "parte" in htxt or "oggetto" in htxt or "titolo" in htxt:
                    col_map["oggetto"] = i
                elif "dettaglio" in htxt or "link" in htxt or "scarica" in htxt:
                    col_map["link"] = i

            idx_data = col_map.get("data", 0)
            idx_oggetto = col_map.get("oggetto", 2)
            idx_link = col_map.get("link", -1)

            page_match = 0
            for row in data_rows:
                cells = row.find_all("td")
                if not cells:
                    continue

                data_pubbl = ""
                oggetto = ""
                link = ""

                # Data
                if idx_data < len(cells):
                    data_pubbl = cells[idx_data].get_text(strip=True)
                else:
                    for cell in cells:
                        dm = re.search(r"\b(\d{2}[/\-]\d{2}[/\-]\d{4})\b", cell.get_text())
                        if dm:
                            data_pubbl = dm.group(1)
                            break

                # Oggetto/Parte
                if idx_oggetto < len(cells):
                    oggetto = cells[idx_oggetto].get_text(strip=True)
                else:
                    # Prendi la cella più lunga
                    for cell in cells:
                        txt = cell.get_text(strip=True)
                        if len(txt) > len(oggetto):
                            oggetto = txt

                # Link
                if idx_link >= 0 and idx_link < len(cells):
                    link_el = cells[idx_link].find("a", href=True)
                    if link_el:
                        link = link_el.get("href", "")
                else:
                    link_el = row.find("a", href=True)
                    if link_el:
                        link = link_el.get("href", "")

                if link and not link.startswith("http"):
                    from urllib.parse import urljoin
                    link = urljoin(base_url, link)

                # Match keyword su oggetto (colonna "Parte")
                if match_keywords(oggetto):
                    results.append({
                        "data": data_pubbl,
                        "oggetto": oggetto,
                        "link": link,
                    })
                    page_match += 1

            log.info(f"TAR - Pagina {page}: {len(data_rows)} righe, {page_match} match")

        # Paginazione
        next_page_found = False
        all_links = soup.find_all("a", href=True)
        for lk in all_links:
            txt = lk.get_text(strip=True).lower()
            href = lk.get("href", "")
            if txt in ("successiva", "next", ">", "»"):
                next_page_found = True
                break
            if re.search(rf"page[=_]{page+1}", href, re.I):
                next_page_found = True
                break

        if not next_page_found:
            break

        page += 1
        time.sleep(0.5)

    log.info(f"TAR Catanzaro: totale {len(results)} atti con keyword")
    return results


# ─────────────────────────────────────────────
# CREAZIONE EXCEL
# ─────────────────────────────────────────────
def create_excel(sheets_data: dict, output_path: str):
    """
    Crea il file Excel con un foglio per portale.
    sheets_data: {sheet_name: [{"data": ..., "oggetto": ..., "link": ...}]}
    """
    wb = Workbook()
    # Rimuovi il foglio default
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Stili
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    link_font = Font(name="Calibri", color="0563C1", underline="single", size=10)
    normal_font = Font(name="Calibri", size=10)
    wrap_align = Alignment(wrap_text=True, vertical="top")
    center_align = Alignment(horizontal="center", vertical="top")

    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    alt_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")

    sheet_order = [
        "ASP Cosenza",
        "ASP Catanzaro",
        "ASP Crotone",
        "ASP Reggio Calabria",
        "ASP Vibo Valentia",
        "Regione Calabria",
        "TAR Catanzaro",
    ]

    for sheet_name in sheet_order:
        rows = sheets_data.get(sheet_name, [])
        ws = wb.create_sheet(title=sheet_name)

        if not rows:
            # Nessun risultato
            ws["A1"] = "Nessun risultato"
            ws["A1"].font = Font(name="Calibri", italic=True, color="808080", size=11)
            ws.column_dimensions["A"].width = 40
            log.info(f"Foglio '{sheet_name}': nessun risultato")
            continue

        # Intestazioni
        headers_cols = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
        for col_idx, h in enumerate(headers_cols, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # Larghezze colonne
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 60
        ws.column_dimensions["C"].width = 40
        ws.column_dimensions["D"].width = 50
        ws.row_dimensions[1].height = 25

        # Dati
        for row_idx, record in enumerate(rows, start=2):
            fill = alt_fill if row_idx % 2 == 0 else None

            # Colonna A: Data
            cell_data = ws.cell(row=row_idx, column=1, value=record.get("data", ""))
            cell_data.font = normal_font
            cell_data.alignment = center_align
            cell_data.border = thin_border
            if fill:
                cell_data.fill = fill

            # Colonna B: Oggetto completo
            oggetto_full = record.get("oggetto", "")
            cell_ogg = ws.cell(row=row_idx, column=2, value=oggetto_full)
            cell_ogg.font = normal_font
            cell_ogg.alignment = wrap_align
            cell_ogg.border = thin_border
            if fill:
                cell_ogg.fill = fill

            # Colonna C: Descrizione breve (primi 150 caratteri)
            desc_breve = oggetto_full[:150] + ("…" if len(oggetto_full) > 150 else "")
            cell_desc = ws.cell(row=row_idx, column=3, value=desc_breve)
            cell_desc.font = normal_font
            cell_desc.alignment = wrap_align
            cell_desc.border = thin_border
            if fill:
                cell_desc.fill = fill

            # Colonna D: Link ipercollegamento
            link_url = record.get("link", "")
            cell_link = ws.cell(row=row_idx, column=4, value=link_url if link_url else "N/D")
            if link_url:
                cell_link.hyperlink = link_url
                cell_link.font = link_font
            else:
                cell_link.font = normal_font
            cell_link.alignment = wrap_align
            cell_link.border = thin_border
            if fill:
                cell_link.fill = fill

            # Altezza riga automatica (approssimazione)
            lines = max(1, len(oggetto_full) // 60 + 1)
            ws.row_dimensions[row_idx].height = max(20, min(lines * 15, 80))

        # Freeze prima riga
        ws.freeze_panes = "A2"

        # Auto-filter
        ws.auto_filter.ref = f"A1:D{len(rows) + 1}"

        log.info(f"Foglio '{sheet_name}': {len(rows)} righe inserite")

    wb.save(output_path)
    log.info(f"File Excel salvato: {output_path}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    log.info("=" * 60)
    log.info("MONITORAGGIO ATTI CALABRIA - 2026-06-05")
    log.info(f"Data dal (ASP + Regione): {DATE_FROM_ASP}")
    log.info(f"Data dal (TAR Catanzaro): {DATE_FROM_TAR}")
    log.info(f"Keyword attive: {len(KEYWORDS)}")
    log.info("=" * 60)

    session = get_session()
    all_sheets = {}

    # ── ASP Portals ──────────────────────────
    for portal in ASP_PORTALS:
        try:
            results = scrape_asp_portal(portal, session)
        except Exception as e:
            log.error(f"{portal['name']}: eccezione non gestita: {e}")
            results = []
        all_sheets[portal["sheet"]] = results

    # ── Regione Calabria ─────────────────────
    try:
        results_regcal = scrape_regione_calabria(session)
    except Exception as e:
        log.error(f"Regione Calabria: eccezione: {e}")
        results_regcal = []
    all_sheets["Regione Calabria"] = results_regcal

    # ── TAR Catanzaro ────────────────────────
    try:
        results_tar = scrape_tar_catanzaro(session)
    except Exception as e:
        log.error(f"TAR Catanzaro: eccezione: {e}")
        results_tar = []
    all_sheets["TAR Catanzaro"] = results_tar

    # ── Sommario ─────────────────────────────
    log.info("=" * 60)
    log.info("SOMMARIO RISULTATI:")
    total = 0
    for sheet, data in all_sheets.items():
        log.info(f"  {sheet}: {len(data)} atti con keyword")
        total += len(data)
    log.info(f"  TOTALE: {total} atti")
    log.info("=" * 60)

    # ── Creazione Excel ──────────────────────
    create_excel(all_sheets, OUTPUT_FILE)
    log.info("COMPLETATO.")
    print(f"\nFile Excel creato: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
