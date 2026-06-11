#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoraggio Atti Calabria
==========================
Scarica e filtra gli atti degli ultimi 2 giorni dai portali:
  - ASP Cosenza, Catanzaro, Crotone, Reggio Calabria, Vibo Valentia (SISR AlboOnline)
  - Regione Calabria provvedimenti
  - TAR Catanzaro (Giustizia Amministrativa)

Output: Monitoraggio_Atti_Calabria_DD-MM-YYYY.xlsx

Dipendenze:
    pip install requests beautifulsoup4 lxml openpyxl
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import re
import time
import os
import sys
import logging
import xml.etree.ElementTree as ET
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("monitoraggio")

# ─────────────────────────────────────────────
# CONFIGURAZIONE
# ─────────────────────────────────────────────
TODAY      = datetime.now()
DATE_FROM  = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_TO    = TODAY.strftime("%Y-%m-%d")
DATE_FROM_YYYYMMDD = (TODAY - timedelta(days=2)).strftime("%Y%m%d")
DATE_TO_YYYYMMDD   = TODAY.strftime("%Y%m%d")

# Per i portali SISR il formato atteso dalla UI è dd/MM/yyyy ma il campo id
# è dataPubblicazioneDal (ISO): usiamo entrambi i formati.
DATE_FROM_IT = (TODAY - timedelta(days=2)).strftime("%d/%m/%Y")
DATE_TO_IT   = TODAY.strftime("%d/%m/%Y")

TAR_DATE_FROM = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")

OUTPUT_DIR  = r"C:\Users\aoedo\kDrive\01_Lavoro\Clienti\Starbene\Atti"
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
)

# Se la cartella di output non esiste, salva nella directory corrente
if not os.path.isdir(OUTPUT_DIR):
    OUTPUT_FILE = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    )
    log.warning(f"Output dir non trovata – salvo in: {OUTPUT_FILE}")

ASP_PORTALS = {
    "ASP Cosenza":       "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":     "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":       "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria":"https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia": "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}

REGIONE_URL = "https://www.regione.calabria.it/provvedimenti-della-regione/"
TAR_URL     = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"

# ─────────────────────────────────────────────
# KEYWORDS
# ─────────────────────────────────────────────
KEYWORDS_PLAIN = [
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

KEYWORDS_UPPER = [k.upper() for k in KEYWORDS_PLAIN]
LIFE_RE = re.compile(r"\bLIFE\b", re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    t = text.upper()
    for kw in KEYWORDS_UPPER:
        if kw in t:
            return True
    return bool(LIFE_RE.search(text))


# ─────────────────────────────────────────────
# HTTP SESSION
# ─────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection":      "keep-alive",
}


def new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# ─────────────────────────────────────────────
# HELPER: estrai tutte le righe da una <table>
# ─────────────────────────────────────────────
DATE_RE = re.compile(r"\d{2}[/\-]\d{2}[/\-]\d{4}|\d{4}[/\-]\d{2}[/\-]\d{2}")


def _extract_rows_from_table(table, base_url: str) -> list[dict]:
    """
    Estrae righe da una tabella HTML.
    Ritorna lista di dict {data, oggetto, link}.
    """
    rows = []
    all_tr = table.find_all("tr")
    # Determina la riga intestazione per capire le colonne
    header_tr = all_tr[0] if all_tr else None
    headers = []
    if header_tr:
        headers = [th.get_text(strip=True).lower() for th in header_tr.find_all(["th", "td"])]

    for tr in all_tr[1:]:
        cells = tr.find_all("td")
        if not cells:
            continue
        texts = [c.get_text(strip=True) for c in cells]

        data_val  = ""
        oggetto   = ""
        link      = ""

        # --- Cerca colonna "data" per indice o contenuto ---
        for idx, h in enumerate(headers):
            if "data" in h and idx < len(texts):
                if DATE_RE.search(texts[idx]):
                    data_val = texts[idx]
                    break

        # --- Fallback: prima cella che sembra una data ---
        if not data_val:
            for t in texts:
                if DATE_RE.search(t):
                    data_val = t
                    break

        # --- Cerca colonna "oggetto" ---
        for idx, h in enumerate(headers):
            if "oggetto" in h and idx < len(texts):
                oggetto = texts[idx]
                a = cells[idx].find("a")
                if a:
                    link = urljoin(base_url, a.get("href", ""))
                break

        # --- Fallback: cella con il testo più lungo (escludendo date/numeri) ---
        if not oggetto:
            best = ("", 0)
            for i, t in enumerate(texts):
                if len(t) > best[1] and not DATE_RE.search(t) and not re.match(r"^\d+$", t):
                    best = (t, len(t))
                    a = cells[i].find("a")
                    link = urljoin(base_url, a.get("href", "")) if a else ""
            oggetto = best[0]

        # --- Link: se ancora vuoto, prima <a> della riga ---
        if not link:
            a = tr.find("a")
            if a:
                link = urljoin(base_url, a.get("href", ""))

        if oggetto or data_val:
            rows.append({"data": data_val, "oggetto": oggetto, "link": link})

    return rows


# ═══════════════════════════════════════════════════════════
# SCRAPER 1: SISR AlboOnline (portali ASP Calabria)
# ═══════════════════════════════════════════════════════════

def _sisr_extract_form_data(soup: BeautifulSoup, url: str) -> tuple[str, dict]:
    """
    Estrae action e tutti i campi hidden dal form JSF.
    Ritorna (action_url, form_dict).
    """
    form = soup.find("form")
    if not form:
        return url, {}

    action = urljoin(url, form.get("action") or url)
    data = {}

    for inp in form.find_all("input"):
        n = inp.get("name")
        v = inp.get("value", "")
        t = inp.get("type", "").lower()
        if n and t not in ("submit", "button", "image", "reset"):
            data[n] = v

    return action, data


def _sisr_set_date_field(soup: BeautifulSoup, form_data: dict, field_id: str,
                         value_iso: str, value_it: str):
    """
    Trova il campo con id=field_id e imposta il suo name nel form_data.
    Prova il valore ISO (YYYY-MM-DD) e italiano (dd/MM/yyyy) in base al
    tipo di input trovato (se type=date → ISO, text → italiano).
    """
    el = soup.find(id=field_id)
    if el and el.get("name"):
        itype = el.get("type", "text").lower()
        form_data[el["name"]] = value_iso if itype == "date" else value_it
        return True

    # Fallback: cerca name che contiene la parte significativa del field_id
    key_hint = field_id.split(":")[-1].lower()
    for inp in soup.find_all("input"):
        name = inp.get("name", "")
        if key_hint in name.lower():
            itype = inp.get("type", "text").lower()
            form_data[name] = value_iso if itype == "date" else value_it
            return True
    return False


def _sisr_find_submit_button(soup: BeautifulSoup) -> tuple[str, str]:
    """Ritorna (name, value) del pulsante di ricerca."""
    for btn in soup.find_all("input", type="submit"):
        v = btn.get("value", "").lower()
        if any(w in v for w in ["cerca", "search", "ricerca", "filtra"]):
            return btn.get("name", ""), btn.get("value", "")
    for btn in soup.find_all("button", type="submit"):
        v = btn.get_text(strip=True).lower()
        if any(w in v for w in ["cerca", "search", "ricerca", "filtra"]):
            return btn.get("name", ""), btn.get_text(strip=True)
    # Primo submit trovato
    btn = soup.find("input", type="submit") or soup.find("button", type="submit")
    if btn:
        return btn.get("name", ""), btn.get("value") or btn.get_text(strip=True)
    return "", ""


def _sisr_find_next_page(soup: BeautifulSoup, session: requests.Session,
                         form_action: str, form_data: dict, base_url: str):
    """
    Cerca il link/bottone 'pagina successiva'.
    Ritorna (response | None, updated_form_data).
    Strategia A: link <a> con testo >, Successiva, Next, …
    Strategia B: JSF PrimeFaces DataTable AJAX pagination.
    """
    # Strategia A – link diretto
    next_a = None
    for candidate in soup.find_all("a"):
        txt = candidate.get_text(strip=True)
        if re.match(r"^(>|›|»|Successiv|Avanti|Next)$", txt, re.I):
            next_a = candidate
            break
        # Paginator numerico: cerca la pagina corrente + 1

    if next_a:
        href = next_a.get("href", "")
        if href and href not in ("#", "javascript:void(0)", "javascript:;", ""):
            url = urljoin(base_url, href)
            try:
                r = session.get(url, timeout=30)
                r.raise_for_status()
                return r, form_data
            except Exception as e:
                log.debug(f"Next-page GET failed: {e}")

    # Strategia B – JSF PrimeFaces DataTable (partial AJAX)
    # Cerca il DataTable e il suo paginatore
    paginator = soup.find(class_=re.compile(r"ui-paginator", re.I))
    if paginator:
        next_btn = paginator.find(class_=re.compile(r"ui-paginator-next", re.I))
        if next_btn and "ui-state-disabled" not in (next_btn.get("class") or []):
            # Trova il DataTable id
            dt = soup.find(attrs={"data-widget": re.compile(r"datatable", re.I)}) or \
                 soup.find(class_=re.compile(r"ui-datatable", re.I))
            dt_id = ""
            if dt:
                dt_id = dt.get("id", "")

            # Calcola first (row offset) dal paginatore
            current_page_el = paginator.find(class_=re.compile(r"ui-paginator-current", re.I))
            rows_per_page = 10  # default
            current_first = int(form_data.get(f"{dt_id}_first", 0)) if dt_id else 0
            next_first = current_first + rows_per_page

            ajax_data = dict(form_data)
            if dt_id:
                ajax_data["javax.faces.partial.ajax"]   = "true"
                ajax_data["javax.faces.source"]         = dt_id
                ajax_data["javax.faces.partial.execute"] = "@all"
                ajax_data["javax.faces.partial.render"]  = dt_id
                ajax_data[dt_id]                         = dt_id
                ajax_data[f"{dt_id}_pagination"]         = "true"
                ajax_data[f"{dt_id}_first"]              = str(next_first)
                ajax_data[f"{dt_id}_rows"]               = str(rows_per_page)
                ajax_data[f"{dt_id}_encodeFeature"]      = "true"

            try:
                r = session.post(form_action, data=ajax_data, timeout=30,
                                 headers={"X-Requested-With": "XMLHttpRequest",
                                          "Faces-Request": "partial/ajax"})
                r.raise_for_status()
                # La risposta AJAX JSF è XML con tag <update id="...">...</update>
                combined_html = r.text  # default: risposta intera
                try:
                    root = ET.fromstring(r.text)
                    for update_el in root.iter("update"):
                        el_id = update_el.get("id", "")
                        if dt_id and dt_id in el_id:
                            combined_html = (update_el.text or "") + \
                                            "".join(ET.tostring(c, encoding="unicode") for c in update_el)
                            break
                except ET.ParseError:
                    pass  # non è XML valido → usa r.text direttamente

                _html = combined_html  # capture per closure

                class _FakeResp:
                    text       = _html
                    status_code = 200

                return _FakeResp(), ajax_data
            except Exception as e:
                log.debug(f"AJAX pagination failed: {e}")

    # Strategia C – submit hidden field "page"
    page_inputs = soup.find_all("input", attrs={"name": re.compile(r"page|pagina", re.I)})
    if page_inputs:
        pi = page_inputs[0]
        try:
            current = int(pi.get("value", "1"))
        except ValueError:
            current = 1
        next_form = dict(form_data)
        next_form[pi["name"]] = str(current + 1)
        try:
            r = session.post(form_action, data=next_form, timeout=30)
            r.raise_for_status()
            return r, next_form
        except Exception as e:
            log.debug(f"Page-input POST failed: {e}")

    return None, form_data


def scrape_asp(portal_name: str, portal_url: str) -> list[dict]:
    """Scrape un portale SISR AlboOnline e ritorna tutti gli atti nel periodo."""
    log.info(f"[{portal_name}] Avvio scraping → {portal_url}")
    session = new_session()
    all_results = []

    try:
        # 1. GET pagina iniziale
        r = session.get(portal_url, timeout=30)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # 2. Estrai form
        action, form_data = _sisr_extract_form_data(soup, portal_url)

        # 3. Imposta date (la funzione sceglie automaticamente ISO vs dd/MM/yyyy)
        _sisr_set_date_field(soup, form_data, "dataPubblicazioneDal", DATE_FROM, DATE_FROM_IT)
        _sisr_set_date_field(soup, form_data, "dataPubblicazioneAl",  DATE_TO,   DATE_TO_IT)

        # 4. Pulsante submit
        btn_name, btn_value = _sisr_find_submit_button(soup)
        if btn_name:
            form_data[btn_name] = btn_value

        log.debug(f"[{portal_name}] POST {action} params={list(form_data.keys())[:8]}…")

        # 5. Esegui ricerca
        r = session.post(action, data=form_data, timeout=30)
        r.raise_for_status()

        # 6. Pagina per pagina
        page_num = 1
        MAX_PAGES = 100

        while page_num <= MAX_PAGES:
            soup = BeautifulSoup(r.text, "html.parser")

            # Cerca la tabella dei risultati
            table = None
            for t in soup.find_all("table"):
                rows = t.find_all("tr")
                if len(rows) >= 2:
                    table = t
                    break

            if not table:
                log.info(f"[{portal_name}] Pagina {page_num}: nessuna tabella trovata, stop.")
                break

            rows_extracted = _extract_rows_from_table(table, portal_url)
            if not rows_extracted:
                log.info(f"[{portal_name}] Pagina {page_num}: tabella vuota, stop.")
                break

            all_results.extend(rows_extracted)
            log.info(f"[{portal_name}] Pagina {page_num}: {len(rows_extracted)} righe (totale {len(all_results)})")

            # Prossima pagina
            r_next, form_data = _sisr_find_next_page(soup, session, action, form_data, portal_url)
            if r_next is None:
                log.info(f"[{portal_name}] Fine paginazione.")
                break

            r = r_next
            page_num += 1
            time.sleep(0.4)

    except requests.exceptions.ConnectionError as e:
        log.error(f"[{portal_name}] Connessione fallita: {e}")
    except requests.exceptions.HTTPError as e:
        log.error(f"[{portal_name}] HTTP error: {e}")
    except Exception as e:
        log.error(f"[{portal_name}] Errore generico: {e}", exc_info=True)

    log.info(f"[{portal_name}] Totale grezzo: {len(all_results)} righe")
    return all_results


# ═══════════════════════════════════════════════════════════
# SCRAPER 2: Regione Calabria Provvedimenti
# ═══════════════════════════════════════════════════════════

def scrape_regione_calabria() -> list[dict]:
    """
    Scarica i provvedimenti della Regione Calabria filtrati per data.
    Il portale usa parametri GET: filter_date_from, filter_date_to, page.
    """
    log.info("[Regione Calabria] Avvio scraping…")
    session = new_session()
    # Alcuni portali WordPress richiedono Referer
    session.headers["Referer"] = REGIONE_URL
    all_results = []

    page = 1
    MAX_PAGES = 100

    while page <= MAX_PAGES:
        params = {
            "filter_date_from": DATE_FROM,
            "filter_date_to":   DATE_TO,
            "paged":            page,
            "page":             page,
        }
        url = f"{REGIONE_URL}?{urlencode(params)}"
        log.debug(f"[Regione Calabria] GET {url}")

        try:
            r = session.get(url, timeout=30)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            # Prova URL alternativo senza "paged"
            if page == 1:
                alt_params = {"filter_date_from": DATE_FROM, "filter_date_to": DATE_TO}
                try:
                    r = session.get(REGIONE_URL, params=alt_params, timeout=30)
                    r.raise_for_status()
                except Exception as e2:
                    log.error(f"[Regione Calabria] Fallito: {e2}")
                    break
            else:
                log.error(f"[Regione Calabria] HTTP {e}")
                break
        except Exception as e:
            log.error(f"[Regione Calabria] Errore: {e}")
            break

        soup = BeautifulSoup(r.text, "html.parser")

        # Cerca tabella principale
        tables = soup.find_all("table")
        rows_this_page = []
        for table in tables:
            extracted = _extract_rows_from_table(table, REGIONE_URL)
            rows_this_page.extend(extracted)

        if not rows_this_page:
            # Cerca struttura alternativa (div/lista)
            items = soup.find_all(class_=re.compile(r"provvedimento|atto|row-item", re.I))
            for item in items:
                text = item.get_text(strip=True)
                a = item.find("a")
                link = urljoin(REGIONE_URL, a.get("href", "")) if a else ""
                date_m = DATE_RE.search(text)
                data_val = date_m.group() if date_m else ""
                rows_this_page.append({"data": data_val, "oggetto": text, "link": link})

        if not rows_this_page:
            log.info(f"[Regione Calabria] Pagina {page}: nessun risultato, stop.")
            break

        all_results.extend(rows_this_page)
        log.info(f"[Regione Calabria] Pagina {page}: {len(rows_this_page)} righe")

        # Controlla se c'è una pagina successiva
        next_link = soup.find("a", rel="next") or \
                    soup.find("a", class_=re.compile(r"next|successiv", re.I)) or \
                    soup.find("a", string=re.compile(r"^(>|›|»|Successiv|Next)$", re.I))
        if not next_link:
            break

        page += 1
        time.sleep(0.4)

    log.info(f"[Regione Calabria] Totale grezzo: {len(all_results)} righe")
    return all_results


# ═══════════════════════════════════════════════════════════
# SCRAPER 3: TAR Catanzaro
# ═══════════════════════════════════════════════════════════

def _tar_search(session: requests.Session, params: dict) -> requests.Response | None:
    """Prova diverse varianti di URL/metodo per il TAR."""
    base = TAR_URL

    # Tentativo 1: GET con parametri
    try:
        r = session.get(base, params=params, timeout=30)
        if r.status_code == 200 and len(r.text) > 500:
            return r
    except Exception as e:
        log.debug(f"TAR GET failed: {e}")

    # Tentativo 2: POST
    try:
        r = session.post(base, data=params, timeout=30)
        if r.status_code == 200 and len(r.text) > 500:
            return r
    except Exception as e:
        log.debug(f"TAR POST failed: {e}")

    # Tentativo 3: URL con formato alternativo
    alt_params = {
        "publishDateFrom": TAR_DATE_FROM,
        "publishDateTo":   DATE_TO,
        "tribunale":       "CZ",
    }
    try:
        r = session.get(base, params=alt_params, timeout=30)
        if r.status_code == 200:
            return r
    except Exception as e:
        log.debug(f"TAR alt GET failed: {e}")

    return None


def scrape_tar_catanzaro() -> list[dict]:
    """
    Scarica i provvedimenti del TAR Catanzaro.
    Filtra la colonna 'Parte' con le keyword (molte parti sono anonimizzate come XXX_OMISSIS_XXX).
    Usa 3 giorni come finestra temporale.
    """
    log.info("[TAR Catanzaro] Avvio scraping…")
    session = new_session()
    all_results = []

    search_params = {
        "publishDateFrom": TAR_DATE_FROM,
        "publishDateTo":   DATE_TO,
        "tribunale":       "CZ",
        "sede":            "CZ",
    }

    page = 1
    MAX_PAGES = 100

    while page <= MAX_PAGES:
        current_params = dict(search_params)
        if page > 1:
            current_params["page"] = page
            current_params["start"] = (page - 1) * 10

        r = _tar_search(session, current_params)
        if r is None:
            log.error("[TAR Catanzaro] Impossibile ottenere risposta dal server.")
            break

        soup = BeautifulSoup(r.text, "html.parser")

        # Cerca tabella risultati
        tables = soup.find_all("table")
        rows_this_page = []

        for table in tables:
            trs = table.find_all("tr")
            if len(trs) < 2:
                continue
            headers = [th.get_text(strip=True).lower() for th in trs[0].find_all(["th", "td"])]

            # Individua indice colonna "Parte" e "Data"
            parte_idx = next((i for i, h in enumerate(headers) if "parte" in h), None)
            data_idx  = next((i for i, h in enumerate(headers) if "data" in h), None)
            num_idx   = next((i for i, h in enumerate(headers) if "numero" in h or "num" == h), None)

            for tr in trs[1:]:
                cells = tr.find_all("td")
                if not cells:
                    continue
                texts = [c.get_text(strip=True) for c in cells]

                parte  = texts[parte_idx] if parte_idx is not None and parte_idx < len(texts) else ""
                data_v = texts[data_idx]  if data_idx  is not None and data_idx  < len(texts) else ""
                numero = texts[num_idx]   if num_idx   is not None and num_idx   < len(texts) else ""

                # Se Parte è omissata, usa il testo completo della riga come oggetto
                oggetto = parte if parte else " | ".join(texts)

                # Link
                link = ""
                for cell in cells:
                    a = cell.find("a")
                    if a:
                        link = urljoin(TAR_URL, a.get("href", ""))
                        break

                rows_this_page.append({
                    "data":    data_v,
                    "oggetto": oggetto,
                    "numero":  numero,
                    "link":    link,
                })

        if not rows_this_page:
            log.info(f"[TAR Catanzaro] Pagina {page}: nessun risultato, stop.")
            break

        all_results.extend(rows_this_page)
        log.info(f"[TAR Catanzaro] Pagina {page}: {len(rows_this_page)} righe")

        # Verifica paginazione
        next_link = soup.find("a", rel="next") or \
                    soup.find("a", class_=re.compile(r"next|successiv", re.I)) or \
                    soup.find("a", string=re.compile(r"^(>|›|»|Successiv|Next|\d+)$", re.I))
        if not next_link:
            # Cerca paginator
            paginator = soup.find(class_=re.compile(r"paginator|pagination", re.I))
            if not paginator:
                break
            disabled_next = paginator.find(class_=re.compile(r"next.*disabled|disabled.*next", re.I))
            if disabled_next:
                break
            active = paginator.find(class_=re.compile(r"active|current", re.I))
            if active:
                next_sibling = active.find_next_sibling()
                if next_sibling:
                    a = next_sibling.find("a")
                    if a:
                        next_link = a

        if not next_link:
            break

        page += 1
        time.sleep(0.4)

    log.info(f"[TAR Catanzaro] Totale grezzo: {len(all_results)} righe")
    return all_results


# ═══════════════════════════════════════════════════════════
# FILTRO KEYWORD
# ═══════════════════════════════════════════════════════════

def filter_results(results: list[dict]) -> list[dict]:
    return [r for r in results if matches_keywords(r.get("oggetto", ""))]


# ═══════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════

HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
LINK_FONT    = Font(name="Calibri", color="0563C1", underline="single", size=10)
DATA_FONT    = Font(name="Calibri", size=10)
ALT_FILL     = PatternFill("solid", fgColor="EBF3FB")
THIN_BORDER  = Border(
    left=Side(style="thin", color="BDD7EE"),
    right=Side(style="thin", color="BDD7EE"),
    top=Side(style="thin", color="BDD7EE"),
    bottom=Side(style="thin", color="BDD7EE"),
)
HEADER_BORDER = Border(
    left=Side(style="medium", color="1F4E79"),
    right=Side(style="medium", color="1F4E79"),
    top=Side(style="medium", color="1F4E79"),
    bottom=Side(style="medium", color="1F4E79"),
)


def _write_sheet(ws, rows: list[dict], portal_name: str):
    """Scrive i dati filtrati in un foglio Excel."""
    ws.freeze_panes = "A2"

    # Intestazioni
    col_labels = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
    for col_idx, label in enumerate(col_labels, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font   = HEADER_FONT
        cell.fill   = HEADER_FILL
        cell.border = HEADER_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 20

    if not rows:
        ws.cell(row=2, column=1, value="Nessun risultato").font = Font(italic=True, color="888888")
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 60
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 40
        return

    for row_idx, item in enumerate(rows, start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")

        # Colonna A – Data
        c_data = ws.cell(row=row_idx, column=1, value=item.get("data", ""))
        c_data.font      = DATA_FONT
        c_data.fill      = fill
        c_data.border    = THIN_BORDER
        c_data.alignment = Alignment(horizontal="center", vertical="top")

        # Colonna B – Oggetto completo
        oggetto = item.get("oggetto", "")
        c_ogg = ws.cell(row=row_idx, column=2, value=oggetto)
        c_ogg.font      = DATA_FONT
        c_ogg.fill      = fill
        c_ogg.border    = THIN_BORDER
        c_ogg.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # Colonna C – Descrizione breve
        breve = (oggetto[:150] + "…") if len(oggetto) > 150 else oggetto
        c_breve = ws.cell(row=row_idx, column=3, value=breve)
        c_breve.font      = DATA_FONT
        c_breve.fill      = fill
        c_breve.border    = THIN_BORDER
        c_breve.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # Colonna D – Link come hyperlink cliccabile
        link = item.get("link", "")
        if link:
            c_link = ws.cell(row=row_idx, column=4, value="Apri atto")
            c_link.hyperlink  = link
            c_link.font       = LINK_FONT
            c_link.fill       = fill
            c_link.border     = THIN_BORDER
            c_link.alignment  = Alignment(horizontal="center", vertical="top")
        else:
            c_link = ws.cell(row=row_idx, column=4, value="—")
            c_link.font      = DATA_FONT
            c_link.fill      = fill
            c_link.border    = THIN_BORDER
            c_link.alignment = Alignment(horizontal="center", vertical="top")

        ws.row_dimensions[row_idx].height = 30 if len(oggetto) > 80 else 20

    # Larghezze colonne
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 18

    # Auto-fit righe (approssimato)
    ws.sheet_view.showGridLines = True


def save_excel(sheets: dict[str, list[dict]]):
    """
    sheets = {nome_foglio: [lista atti filtrati], …}
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuovi foglio vuoto default

    SHEET_ORDER = [
        "ASP Cosenza",
        "ASP Catanzaro",
        "ASP Crotone",
        "ASP Reggio Calabria",
        "ASP Vibo Valentia",
        "Regione Calabria",
        "TAR Catanzaro",
    ]

    for name in SHEET_ORDER:
        data = sheets.get(name, [])
        ws = wb.create_sheet(title=name)
        _write_sheet(ws, data, name)
        count = len(data)
        log.info(f"  Foglio '{name}': {count} match")

    # Foglio riepilogo
    ws_sum = wb.create_sheet(title="RIEPILOGO", index=0)
    ws_sum["A1"] = "Monitoraggio Atti Calabria"
    ws_sum["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws_sum["A2"] = f"Generato il: {TODAY.strftime('%d/%m/%Y %H:%M')}"
    ws_sum["A3"] = f"Periodo: {DATE_FROM_IT} – {DATE_TO_IT}"
    ws_sum["A4"] = f"TAR: ultimi 3 gg ({(TODAY-timedelta(days=3)).strftime('%d/%m/%Y')} – {DATE_TO_IT})"
    ws_sum["A5"] = ""
    ws_sum["A6"] = "Portale"
    ws_sum["B6"] = "N° match"
    ws_sum["A6"].font = HEADER_FONT
    ws_sum["A6"].fill = HEADER_FILL
    ws_sum["B6"].font = HEADER_FONT
    ws_sum["B6"].fill = HEADER_FILL

    for i, name in enumerate(SHEET_ORDER, start=7):
        ws_sum[f"A{i}"] = name
        ws_sum[f"B{i}"] = len(sheets.get(name, []))
        ws_sum[f"B{i}"].alignment = Alignment(horizontal="center")

    ws_sum.column_dimensions["A"].width = 25
    ws_sum.column_dimensions["B"].width = 12

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    wb.save(OUTPUT_FILE)
    log.info(f"\n✓ File salvato: {OUTPUT_FILE}")


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    log.info("=" * 60)
    log.info("  MONITORAGGIO ATTI CALABRIA")
    log.info(f"  Periodo: {DATE_FROM_IT} – {DATE_TO_IT}")
    log.info("=" * 60)

    sheets: dict[str, list[dict]] = {}

    # 1. Portali ASP
    for portal_name, portal_url in ASP_PORTALS.items():
        raw = scrape_asp(portal_name, portal_url)
        sheets[portal_name] = filter_results(raw)
        time.sleep(0.5)

    # 2. Regione Calabria
    raw = scrape_regione_calabria()
    sheets["Regione Calabria"] = filter_results(raw)
    time.sleep(0.5)

    # 3. TAR Catanzaro
    raw = scrape_tar_catanzaro()
    # Per il TAR filtriamo su "oggetto" che può contenere la colonna "Parte"
    # Le righe con XXX_OMISSIS_XXX non matcheranno per keyword → saranno escluse
    sheets["TAR Catanzaro"] = filter_results(raw)

    # 4. Salva Excel
    log.info("\nSalvataggio Excel…")
    save_excel(sheets)

    # 5. Riepilogo finale
    log.info("\n" + "=" * 60)
    log.info("  RIEPILOGO MATCH")
    log.info("=" * 60)
    total = 0
    for name, rows in sheets.items():
        log.info(f"  {name:<25} {len(rows):>4} match")
        total += len(rows)
    log.info(f"  {'TOTALE':<25} {total:>4}")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
