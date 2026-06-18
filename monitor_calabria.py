#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria - versione finale
ASP Cosenza/Catanzaro/Crotone/RC/VV, Regione Calabria, TAR Catanzaro
"""

import sys
import time
import re
import json
import traceback
from datetime import date, timedelta
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── date ──────────────────────────────────────────────────────────────────────
TODAY       = date.today()
DATE_2DAYS  = TODAY - timedelta(days=2)
DATE_3DAYS  = TODAY - timedelta(days=3)
DATE_STR_2  = DATE_2DAYS.strftime("%Y-%m-%d")
DATE_STR_3  = DATE_3DAYS.strftime("%Y-%m-%d")
DATE_TODAY  = TODAY.strftime("%Y-%m-%d")

# ── keywords ──────────────────────────────────────────────────────────────────
KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento",
    "San Teodoro", "Savelli Hospital", "Starbene",
    "Verifica requisiti", "Villa San Giuseppe", "Villa del Rosario",
]
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def matches_keyword(text: str) -> bool:
    if not text:
        return False
    tl = text.lower()
    for kw in KEYWORDS:
        if kw.lower() in tl:
            return True
    if LIFE_RE.search(text):
        return True
    return False

# ── HTTP helpers ──────────────────────────────────────────────────────────────
BASE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(BASE_HEADERS)
    return s

def _retry(fn, retries=3, base_delay=3):
    for attempt in range(retries):
        try:
            r = fn()
            r.raise_for_status()
            return r
        except requests.HTTPError as e:
            code = e.response.status_code
            print(f"  [WARN] HTTP {code} (tentativo {attempt+1})")
            if attempt < retries - 1:
                time.sleep(base_delay * (attempt + 1))
        except Exception as e:
            print(f"  [WARN] {e} (tentativo {attempt+1})")
            if attempt < retries - 1:
                time.sleep(base_delay)
    return None

def safe_get(session, url, params=None, extra_headers=None):
    h = dict(extra_headers or {})
    return _retry(lambda: session.get(url, params=params, headers=h,
                                      timeout=30, verify=False))

def safe_post(session, url, data=None, json_data=None, extra_headers=None):
    h = dict(extra_headers or {})
    return _retry(lambda: session.post(url, data=data, json=json_data,
                                       headers=h, timeout=30, verify=False))


# ══════════════════════════════════════════════════════════════════════════════
# ASP PORTALS (SISR)
# ══════════════════════════════════════════════════════════════════════════════

ASP_PORTALS = {
    "ASP Cosenza":         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}

def parse_sisr_table(soup, base_url):
    rows = []
    best_table, best_score = None, 0
    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if len(trs) < 2:
            continue
        header_text = " ".join(th.get_text(strip=True).lower()
                                for th in trs[0].find_all(["th", "td"]))
        score = sum(kw in header_text for kw in ["data", "oggetto", "numero", "tipo"])
        if score > best_score:
            best_score, best_table = score, table

    if best_table is None:
        return rows

    trs = best_table.find_all("tr")
    headers = [th.get_text(strip=True).lower() for th in trs[0].find_all(["th","td"])]

    idx_data = next((i for i, h in enumerate(headers)
                     if any(x in h for x in ["data pubbl", "data di pub", "data_pub", "pubblicaz"])), None)
    idx_ogg  = next((i for i, h in enumerate(headers)
                     if any(x in h for x in ["oggetto", "titolo"])), None)
    if idx_data is None:
        idx_data = next((i for i, h in enumerate(headers) if "data" in h), 0)

    for tr in trs[1:]:
        tds = tr.find_all("td")
        if not tds or len(tds) < 2:
            continue
        row_data = [td.get_text(" ", strip=True) for td in tds]

        link = ""
        for td in tds:
            a = td.find("a", href=True)
            if a and not a["href"].startswith("#"):
                href = a["href"]
                link = href if href.startswith("http") else urljoin(base_url, href)
                break

        data_val = (row_data[idx_data] if idx_data is not None and idx_data < len(row_data)
                    else (row_data[0] if row_data else ""))

        if idx_ogg is not None and idx_ogg < len(row_data):
            ogg = row_data[idx_ogg]
        else:
            # Prendi il campo testuale più lungo (escludi "→" e date)
            candidates = [v for v in row_data if len(v) > 10 and not re.match(r'^\d', v) and v != "→"]
            ogg = max(candidates, key=len) if candidates else (max(row_data, key=len) if row_data else "")

        rows.append({"data": data_val, "oggetto": ogg, "link": link})
    return rows

def fetch_asp_portal(name, base_url):
    print(f"\n>>> {name}: {base_url}")
    session = make_session()

    resp = safe_get(session, base_url,
                    extra_headers={"Referer": "https://www.regione.calabria.it/"})
    if resp is None:
        print(f"  [ERR] Non raggiungibile (HTTP 503) - il server SISR blocca IP cloud")
        return []

    soup = BeautifulSoup(resp.text, "lxml")
    form = soup.find("form")
    if not form:
        print("  [WARN] Nessun form trovato")
        return []

    action  = form.get("action", "")
    method  = form.get("method", "get").lower()
    submit_url = urljoin(base_url, action) if action else base_url

    hidden = {i.get("name",""): i.get("value","")
              for i in form.find_all("input", type="hidden") if i.get("name")}

    payload = dict(hidden)
    payload["dataPubblicazioneDal"] = DATE_STR_2
    payload["dataPubblicazioneAl"]  = DATE_TODAY

    if method == "post":
        resp2 = safe_post(session, submit_url, data=payload,
                          extra_headers={"Referer": base_url})
    else:
        resp2 = safe_get(session, submit_url, params=payload,
                         extra_headers={"Referer": base_url})

    if resp2 is None:
        return []

    all_rows = []
    soup2 = BeautifulSoup(resp2.text, "lxml")
    page_rows = parse_sisr_table(soup2, base_url)
    print(f"  Pagina 1: {len(page_rows)} atti")
    all_rows.extend(page_rows)

    page_num = 2
    while page_num <= 100:
        next_url = None
        for a in soup2.find_all("a", href=True):
            txt = a.get_text(strip=True).lower()
            if txt in ["successivo", "next", ">", "»"] or txt == str(page_num):
                href = a["href"]
                next_url = href if href.startswith("http") else urljoin(base_url, href)
                break
        if not next_url:
            break
        r = safe_get(session, next_url, extra_headers={"Referer": resp2.url})
        if not r:
            break
        soup2 = BeautifulSoup(r.text, "lxml")
        pr = parse_sisr_table(soup2, base_url)
        if not pr:
            break
        print(f"  Pagina {page_num}: {len(pr)} atti")
        all_rows.extend(pr)
        page_num += 1

    print(f"  Totale: {len(all_rows)} | Match: ", end="")
    filtered = [r for r in all_rows if matches_keyword(r["oggetto"])]
    print(len(filtered))
    return filtered


# ══════════════════════════════════════════════════════════════════════════════
# REGIONE CALABRIA
# ══════════════════════════════════════════════════════════════════════════════
# Tabella: Tipologia | Data Repertoriazione | N | Dipartimento | Oggetto | Dettaglio
# Le righe hanno solo 5 TD (la colonna "N" è spesso vuota/assente).
# Mapping effettivo: [0]=Tipologia [1]=Data [2]=Dipartimento [3]=Oggetto [4]=→link

RC_BASE = "https://www.regione.calabria.it/provvedimenti-della-regione/"

def fetch_regione_calabria():
    print(f"\n>>> Regione Calabria: {RC_BASE}")
    session = make_session()

    def parse_rc_table(soup):
        rows = []
        table = soup.find("table")
        if not table:
            return rows
        trs = table.find_all("tr")
        if len(trs) < 2:
            return rows

        for tr in trs[1:]:
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue

            data_val = tds[1].get_text(" ", strip=True)   # Data Repertoriazione

            # Oggetto è l'ultimo campo lungo prima del "→"
            # Con 5 TDs: [0]Tipologia [1]Data [2]Dipartimento [3]Oggetto [4]→
            # Con 6 TDs: [0]Tipologia [1]Data [2]N [3]Dipartimento [4]Oggetto [5]→
            n_tds = len(tds)
            ogg_idx = n_tds - 2   # Penultimo (prima del link →)
            ogg_val = tds[ogg_idx].get_text(" ", strip=True)

            # Link: nell'ultimo TD con <a>
            link = ""
            for td in reversed(tds):
                a = td.find("a", href=True)
                if a and a["href"].startswith("http"):
                    link = a["href"]
                    break

            rows.append({"data": data_val, "oggetto": ogg_val, "link": link})
        return rows

    all_rows  = []
    seen_keys = set()
    page_num  = 1

    while page_num <= 300:
        if page_num == 1:
            url = f"{RC_BASE}?filter_date_from={DATE_STR_2}"
        else:
            url = f"{RC_BASE}?paged={page_num}&filter_date_from={DATE_STR_2}"

        resp = safe_get(session, url, extra_headers={"Referer": RC_BASE})
        if not resp:
            break

        soup = BeautifulSoup(resp.text, "lxml")
        page_rows = parse_rc_table(soup)
        if not page_rows:
            break

        new_rows = []
        for r in page_rows:
            key = r["link"] or r["oggetto"]
            if key not in seen_keys:
                seen_keys.add(key)
                new_rows.append(r)

        if not new_rows:
            break

        print(f"  Pagina {page_num}: {len(page_rows)} atti ({len(new_rows)} nuovi)")
        all_rows.extend(new_rows)

        # Prossima pagina: cerca link "Pagina successiva" o paged=N+1
        has_next = bool(
            soup.find("a", string=re.compile(r"successiva", re.I)) or
            soup.find("a", rel="next") or
            soup.find("a", href=re.compile(rf"paged={page_num+1}"))
        )
        if not has_next:
            break
        page_num += 1

    print(f"  Totale: {len(all_rows)} | Match: ", end="")
    filtered = [r for r in all_rows if matches_keyword(r["oggetto"])]
    print(len(filtered))
    return filtered


# ══════════════════════════════════════════════════════════════════════════════
# TAR CATANZARO  (Liferay portlet + DataTables AJAX)
# ══════════════════════════════════════════════════════════════════════════════
# L'endpoint DataTables è p_p_lifecycle=2 / p_p_resource_id=/administrative-acts/search/results
# La ricerca avviene tramite POST JSON con campo additionalInfo (JSON stringa).

TAR_BASE = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"

def fetch_tar_catanzaro():
    print(f"\n>>> TAR Catanzaro: {TAR_BASE}")
    session = make_session()

    # Step 1: GET pagina per ottenere p_p_id e p_auth dalla action del form
    resp0 = safe_get(session, TAR_BASE)
    if not resp0:
        print("  [ERR] Portale TAR non raggiungibile")
        return []

    soup0 = BeautifulSoup(resp0.text, "lxml")

    # Trova form con publishDateFrom
    target_form = None
    for form in soup0.find_all("form"):
        if form.find("input", attrs={"name": re.compile(r"publishDateFrom")}):
            target_form = form
            break

    if not target_form:
        print("  [ERR] Form TAR non trovato")
        return []

    action_form = target_form.get("action", "")
    prefix = ""
    for inp in target_form.find_all("input"):
        n = inp.get("name", "")
        if "publishDateFrom" in n:
            prefix = n.replace("publishDateFrom", "")
            break

    # Step 2: POST form per ottenere la pagina con lo script AJAX
    post_data = {}
    for inp in target_form.find_all("input"):
        n, v = inp.get("name",""), inp.get("value","")
        if n:
            post_data[n] = v
    for sel in target_form.find_all("select"):
        n = sel.get("name","")
        if n:
            opt = sel.find("option", selected=True)
            post_data[n] = opt.get("value","") if opt else ""

    post_data[f"{prefix}publishDateFrom"] = DATE_STR_3
    post_data[f"{prefix}publishDateTo"]   = DATE_TODAY

    resp_post = safe_post(session, action_form, data=post_data, extra_headers={
        "Referer": TAR_BASE,
        "Origin": "https://www.giustizia-amministrativa.it",
        "Content-Type": "application/x-www-form-urlencoded",
    })
    if not resp_post:
        print("  [ERR] POST TAR fallito")
        return []

    soup_post = BeautifulSoup(resp_post.text, "lxml")

    # Estrai searchURL e additionalInfo dallo script
    resource_url  = None
    additional_info_template = None
    for s in soup_post.find_all("script", string=True):
        txt = s.string
        if "searchURL" in txt and "additionalInfo" in txt:
            m_url = re.search(r'var\s+searchURL\s*=\s*"([^"]+)"', txt)
            m_add = re.search(r"d\.additionalInfo\s*=\s*'([^']+)'", txt)
            if m_url:
                resource_url = m_url.group(1)
            if m_add:
                additional_info_template = m_add.group(1)
            break

    if not resource_url:
        print("  [ERR] Resource URL TAR non trovato nello script")
        return []

    print(f"  Resource URL: {resource_url[:100]}")
    print(f"  additionalInfo: {additional_info_template[:100] if additional_info_template else 'N/A'}")

    # Step 3: Chiama l'endpoint DataTables iterando le pagine
    all_rows   = []
    page_start = 0
    page_size  = 100  # Chiedi 100 per volta per ridurre le pagine
    draw_num   = 1

    while True:
        # Costruisce additionalInfo con la data corretta (già impostata dal POST, ma ricostruiamo)
        add_info = {
            "schema": "TAR_CATANZARO",
            "type": None,
            "year": "",
            "number": "",
            "hearingDateFrom": None,
            "hearingDateTo": None,
            "publishDateFrom": DATE_STR_3,
            "publishDateTo": DATE_TODAY,
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
            "queryString": None,
        }

        payload = {
            "draw": draw_num,
            "start": page_start,
            "length": page_size,
            "search": {"value": "", "regex": False},
            "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
            "additionalInfo": json.dumps(add_info, ensure_ascii=False),
        }

        resp_ajax = safe_post(
            session,
            resource_url,
            json_data=payload,
            extra_headers={
                "Referer": resp_post.url,
                "Content-Type": "application/json",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
                "Origin": "https://www.giustizia-amministrativa.it",
            }
        )

        if resp_ajax is None:
            print(f"  [ERR] AJAX call fallita (start={page_start})")
            break

        try:
            data_json = resp_ajax.json()
        except Exception as e:
            print(f"  [ERR] Risposta non è JSON: {e}")
            print(f"  Response: {resp_ajax.text[:300]}")
            break

        if data_json.get("error"):
            print(f"  [ERR] Server error: {data_json.get('error')}")
            break

        records = data_json.get("data", [])
        total   = data_json.get("recordsTotal", 0)

        print(f"  start={page_start}: {len(records)} record (totale={total})")

        for rec in records:
            nrg          = rec.get("nrgFascicolo", "")
            parte        = rec.get("parte", "")
            data_pub     = rec.get("dataPubblicazione", "")
            tipo_prov    = rec.get("tipoProvvedimento", "")
            num_prov     = rec.get("numProvvedimento", "")
            nome_file    = rec.get("nomeFile", "")

            # Link al documento
            link = ""
            if nome_file:
                link = (f"https://mdp.giustizia-amministrativa.it/visualizza/"
                        f"?nodeRef=&schema=tar_cz&nrg={nrg}&nomeFile={nome_file}&subDir=Provvedimenti")
            elif nrg:
                link = f"{TAR_BASE}?nrg={nrg}"

            # Oggetto per matching: parte + tipo + numero
            oggetto_match = f"{tipo_prov} {parte} {num_prov}".strip()

            all_rows.append({
                "data":          data_pub,
                "oggetto":       parte,             # Mostrato nel foglio: colonna Parte
                "oggetto_match": oggetto_match,     # Usato per keyword match
                "link":          link,
            })

        page_start += len(records)
        draw_num   += 1

        if page_start >= total or not records:
            break

        time.sleep(0.5)

    print(f"  Totale recuperati: {len(all_rows)}")
    filtered = [r for r in all_rows
                if "XXX_OMISSIS_XXX" not in r.get("oggetto","").upper()
                and matches_keyword(r.get("oggetto_match", r.get("oggetto","")))]
    print(f"  Match keyword: {len(filtered)}")
    return filtered


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

HYPERLINK_FONT = Font(color="0000FF", underline="single", name="Calibri", size=11)
HEADER_FONT    = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
HEADER_FILL    = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
ALT_FILL       = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
NORMAL_FONT    = Font(name="Calibri", size=11)
ERR_FONT       = Font(name="Calibri", size=11, color="C00000")

def write_sheet(ws, rows, error_msg=None):
    cols       = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
    col_widths = [16, 65, 55, 40]

    for ci, (col, w) in enumerate(zip(cols, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22

    if error_msg:
        c = ws.cell(row=2, column=1, value=error_msg)
        ws.merge_cells("A2:D2")
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        c.font = ERR_FONT
        ws.row_dimensions[2].height = 36
        return

    if not rows:
        c = ws.cell(row=2, column=1, value="Nessun risultato per il periodo selezionato")
        ws.merge_cells("A2:D2")
        c.alignment = Alignment(horizontal="center")
        c.font = NORMAL_FONT
        return

    for ri, row in enumerate(rows, 2):
        fill = ALT_FILL if ri % 2 == 0 else None

        def sc(col_i, value, wrap=False, center=False):
            c = ws.cell(row=ri, column=col_i, value=value)
            c.font = NORMAL_FONT
            c.alignment = Alignment(wrap_text=wrap,
                                    horizontal="center" if center else "left",
                                    vertical="top")
            if fill:
                c.fill = fill
            return c

        sc(1, row.get("data", ""), center=True)
        sc(2, row.get("oggetto", ""), wrap=True)
        oggetto = row.get("oggetto", "")
        sc(3, oggetto[:150] if oggetto else "", wrap=True)

        link = row.get("link", "")
        c4 = ws.cell(row=ri, column=4)
        if link:
            c4.value     = "Apri atto"
            c4.hyperlink = link
            c4.font      = HYPERLINK_FONT
        else:
            c4.value = ""
            c4.font  = NORMAL_FONT
        c4.alignment = Alignment(horizontal="center", vertical="top")
        if fill:
            c4.fill = fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{len(rows)+1}"


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print(f"Monitoraggio Atti Calabria — {TODAY.strftime('%d/%m/%Y')}")
    print(f"ASP / RC periodo : {DATE_STR_2} → {DATE_TODAY}")
    print(f"TAR periodo      : {DATE_STR_3} → {DATE_TODAY}")
    print("=" * 70)

    results = {}
    errors  = {}

    for name, url in ASP_PORTALS.items():
        try:
            results[name] = fetch_asp_portal(name, url)
        except Exception as e:
            traceback.print_exc()
            results[name] = []
            errors[name]  = str(e)

    try:
        results["Regione Calabria"] = fetch_regione_calabria()
    except Exception as e:
        traceback.print_exc()
        results["Regione Calabria"] = []
        errors["Regione Calabria"]  = str(e)

    try:
        results["TAR Catanzaro"] = fetch_tar_catanzaro()
    except Exception as e:
        traceback.print_exc()
        results["TAR Catanzaro"] = []
        errors["TAR Catanzaro"]  = str(e)

    # ── Excel ──────────────────────────────────────────────────────────────
    fname = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    fpath = f"/home/user/NS/{fname}"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_order = [
        "ASP Cosenza", "ASP Catanzaro", "ASP Crotone",
        "ASP Reggio Calabria", "ASP Vibo Valentia",
        "Regione Calabria", "TAR Catanzaro",
    ]

    for sname in sheet_order:
        ws = wb.create_sheet(title=sname)
        rows = results.get(sname, [])
        err  = errors.get(sname)
        if err and not rows:
            errmsg = (f"⚠ Portale non raggiungibile dall'ambiente cloud (HTTP 503).\n"
                      f"URL: {ASP_PORTALS.get(sname, 'N/A')}\n"
                      f"I portali SISR Regione Calabria bloccano le richieste da IP cloud/datacenter. "
                      f"Eseguire lo script dalla rete locale per ottenere i dati.")
            write_sheet(ws, rows, error_msg=errmsg)
        else:
            write_sheet(ws, rows)
        print(f"  Foglio '{sname}': {len(rows)} match{' [ERR CONN]' if (err and not rows) else ''}")

    wb.save(fpath)
    print(f"\n{'='*70}")
    print(f"File: {fpath}")
    print(f"{'='*70}")
    print("\nRIEPILOGO:")
    total_match = 0
    for sname in sheet_order:
        n = len(results.get(sname, []))
        e = " ⚠ ERRORE CONNESSIONE" if sname in errors and not n else ""
        print(f"  {sname:<25} {n} match{e}")
        total_match += n
    print(f"  {'TOTALE':<25} {total_match}")

    return fpath, results, errors


if __name__ == "__main__":
    fpath, results, errors = main()
