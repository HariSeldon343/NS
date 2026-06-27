#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria – FINAL
- ASP SISR portals:    inaccessibili (TLS handshake timeout – server regionale bloccato da IP cloud)
- Regione Calabria:    GET paginato, filtro client-side per data
- TAR Catanzaro:       POST DataTables API JSON
"""

import re, json
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── Dates ────────────────────────────────────────────────────────────────────
TODAY      = date(2026, 6, 27)
DATE_2D    = (TODAY - timedelta(days=2)).isoformat()   # 2026-06-25
DATE_3D    = (TODAY - timedelta(days=3)).isoformat()   # 2026-06-24
DATE_TODAY = TODAY.isoformat()

def parse_it_date(s: str):
    """Parse DD/MM/YYYY Italian date, return date object or None."""
    try:
        d, m, y = s.strip().split("/")
        return date(int(y), int(m), int(d))
    except Exception:
        return None

# ─── Keywords ────────────────────────────────────────────────────────────────
KEYWORDS_RAW = [
    "ADI",
    "Assistenza domiciliare",
    "ANMIC",
    "Accreditamento",
    "Aumento di budget",
    "Autismo",
    "Autorizzazione all.esercizio",
    "Autorizzazione alla realizzazione",
    "Autorizzazioni",
    "Budget",
    "Casa Giardino",
    "Centro San Giuseppe",
    "Centro salute e benessere",
    "Fabbisogni LEA",
    "Fisiolab",
    "Fisioterapia",
    r"\bLIFE\b",            # word boundary
    "Parere commissione",
    "Presa d.atto verifica",
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

KW_PATTERN = re.compile(
    "|".join(k if k.startswith(r"\b") else re.escape(k) for k in KEYWORDS_RAW),
    re.IGNORECASE,
)

def matches(text: str) -> bool:
    return bool(KW_PATTERN.search(text or ""))

# ─── HTTP session ─────────────────────────────────────────────────────────────
SESSION = requests.Session()
SESSION.verify = "/root/.ccr/ca-bundle.crt"
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9",
})

# ─── Output ───────────────────────────────────────────────────────────────────
OUT_FILENAME = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
OUT_PATH     = Path("/home/user/NS") / OUT_FILENAME

# ─── Excel helpers ────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
LINK_FONT   = Font(color="0563C1", underline="single", size=10)
BODY_FONT   = Font(size=10)
ITALIC_FONT = Font(italic=True, color="AA0000", size=9)
WRAP        = Alignment(wrap_text=True, vertical="top")
COLUMNS     = ["Data", "Oggetto", "Descrizione breve", "Link"]
COL_WIDTHS  = [14, 55, 52, 12]

def create_sheet(wb: openpyxl.Workbook, name: str, rows: list, note: str = "") -> None:
    ws = wb.create_sheet(title=name[:31])
    for ci, (col, w) in enumerate(zip(COLUMNS, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 18
    ws.freeze_panes = "A2"

    if note:
        ws.cell(row=2, column=1, value=f"NOTA: {note}").font = ITALIC_FONT
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
        return

    if not rows:
        ws.cell(row=2, column=1, value="Nessun risultato").font = BODY_FONT
        return

    for ri, row in enumerate(rows, 2):
        data    = row.get("data", "")
        oggetto = row.get("oggetto", "")
        descr   = (oggetto[:150] + "…") if len(oggetto) > 150 else oggetto
        link    = row.get("link", "")

        ws.cell(row=ri, column=1, value=data).font = BODY_FONT
        c2 = ws.cell(row=ri, column=2, value=oggetto)
        c2.font = BODY_FONT; c2.alignment = WRAP
        c3 = ws.cell(row=ri, column=3, value=descr)
        c3.font = BODY_FONT; c3.alignment = WRAP

        lc = ws.cell(row=ri, column=4, value="Apri" if link else "—")
        if link:
            lc.hyperlink = link
            lc.font = LINK_FONT
        else:
            lc.font = BODY_FONT
        ws.row_dimensions[ri].height = 45


# ══════════════════════════════════════════════════════════════════════════════
# ASP SISR portals
# ══════════════════════════════════════════════════════════════════════════════

ASP_PORTALS = [
    ("ASP Cosenza",         "https://online-aspco.sisr.regione.calabria.it"),
    ("ASP Catanzaro",       "https://online-aspcz.sisr.regione.calabria.it"),
    ("ASP Crotone",         "https://online-aspkr.sisr.regione.calabria.it"),
    ("ASP Reggio Calabria", "https://online-asprc.sisr.regione.calabria.it"),
    ("ASP Vibo Valentia",   "https://online-aspvv.sisr.regione.calabria.it"),
]

def scrape_asp(name: str, base: str) -> tuple:
    """Attempt ASP scrape. Returns (list_of_rows, error_note)."""
    url = f"{base}/AlboOnline/ricercaAlbo"
    print(f"[{name}] → {url}")
    try:
        r = SESSION.get(url, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")

        form = soup.find("form")
        if not form:
            return [], "Form non trovato nella pagina"

        form_data = {}
        for inp in form.find_all("input"):
            n = inp.get("name"); v = inp.get("value", "")
            if n:
                form_data[n] = v

        # Set date fields
        for k in ("dataPubblicazioneDal", "dataDal", "dataInizio"):
            if k in form_data:
                form_data[k] = DATE_2D; break
        else:
            form_data["dataPubblicazioneDal"] = DATE_2D

        for k in ("dataPubblicazioneAl", "dataAl", "dataFine"):
            if k in form_data:
                form_data[k] = DATE_TODAY; break

        action = form.get("action") or url
        if not action.startswith("http"):
            action = base + "/" + action.lstrip("/")
        method = (form.get("method") or "POST").upper()

        all_rows = []
        pg = 1
        cur_url = url

        while True:
            if method == "POST":
                r2 = SESSION.post(action, data=form_data, timeout=15)
            else:
                r2 = SESSION.get(action, params=form_data, timeout=15)

            soup2 = BeautifulSoup(r2.text, "lxml")
            rows = _extract_asp_rows(soup2, base)
            all_rows.extend(rows)
            print(f"  page {pg}: {len(rows)} rows")

            next_a = None
            for a in soup2.find_all("a"):
                t = a.get_text(strip=True).lower()
                if t in ("successivo", "next", "avanti", ">", "»"):
                    next_a = a.get("href"); break
            if not next_a:
                break

            if not next_a.startswith("http"):
                next_a = urljoin(cur_url, next_a)
            cur_url = next_a
            r2 = SESSION.get(next_a, timeout=15)
            pg += 1
            if pg > 100:
                break

        return all_rows, ""

    except requests.exceptions.SSLError as e:
        msg = (f"Server SISR ({base.split('/')[2]}) non raggiungibile da ambiente cloud "
               f"(TLS handshake timeout). Verificare da rete italiana. Errore: {str(e)[:80]}")
        print(f"  [SSL] {msg[:100]}")
        return [], msg
    except requests.exceptions.Timeout:
        msg = f"Timeout connessione a {base.split('/')[2]}"
        print(f"  [TIMEOUT]")
        return [], msg
    except requests.exceptions.ConnectionError as e:
        msg = f"Connessione rifiutata: {str(e)[:80]}"
        print(f"  [CONN] {msg[:80]}")
        return [], msg
    except Exception as e:
        msg = f"{type(e).__name__}: {str(e)[:80]}"
        print(f"  [ERR] {msg}")
        return [], msg


def _extract_asp_rows(soup, base):
    rows = []
    table = soup.find("table")
    if not table:
        return rows
    trs = table.find_all("tr")
    if not trs:
        return rows
    headers = [c.get_text(strip=True).lower() for c in trs[0].find_all(["th","td"])]
    ogg_i  = next((i for i,h in enumerate(headers) if "oggetto" in h), None)
    data_i = next((i for i,h in enumerate(headers) if "data" in h), None)

    for tr in trs[1:]:
        cells = [c.get_text(strip=True) for c in tr.find_all(["td","th"])]
        if not cells: continue
        link_a = tr.find("a")
        link = ""
        if link_a and link_a.get("href"):
            h = link_a["href"]
            link = h if h.startswith("http") else urljoin(base, h)
        ogg = cells[ogg_i] if ogg_i is not None and ogg_i < len(cells) else max(cells, key=len)
        dat = cells[data_i] if data_i is not None and data_i < len(cells) else ""
        if ogg:
            rows.append({"data": dat, "oggetto": ogg, "link": link})
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# REGIONE CALABRIA – paginate GET, client-side date filter
# ══════════════════════════════════════════════════════════════════════════════

def scrape_regione_calabria() -> list:
    """
    Paginates through /provvedimenti-della-regione/ newest-first.
    Stops when rows older than DATE_2D appear.
    Returns rows where date >= DATE_2D.
    """
    BASE     = "https://www.regione.calabria.it/provvedimenti-della-regione/"
    CUTOFF   = date.fromisoformat(DATE_2D)
    print(f"\n[Regione Calabria] Paginating from {DATE_2D} (newest first)")

    all_rows = []

    for pg in range(1, 500):
        url = BASE if pg == 1 else f"{BASE}?paged={pg}"
        try:
            r = SESSION.get(url, timeout=20)
        except Exception as e:
            print(f"  Page {pg} error: {e}"); break

        soup = BeautifulSoup(r.text, "lxml")
        table = soup.find("table")
        if not table:
            print(f"  Page {pg}: no table"); break

        trs = table.find_all("tr")[1:]  # skip header
        if not trs:
            print(f"  Page {pg}: empty"); break

        # Table data cols (N column NOT rendered in data rows):
        # 0=Tipologia, 1=DataRepertoriazione, 2=Dipartimento, 3=Oggetto, 4=Dettaglio
        page_rows = []
        stop = False
        for tr in trs:
            cells = [c.get_text(strip=True) for c in tr.find_all("td")]
            if len(cells) < 4:
                continue

            date_str  = cells[1]
            row_date  = parse_it_date(date_str)
            oggetto   = cells[3]
            link_a    = tr.find("a")
            link      = link_a.get("href", "") if link_a else ""
            if link and not link.startswith("http"):
                link = "https://www.regione.calabria.it" + link

            if row_date is not None and row_date < CUTOFF:
                stop = True
                break

            page_rows.append({"data": date_str, "oggetto": oggetto, "link": link})

        all_rows.extend(page_rows)
        print(f"  Page {pg}: {len(page_rows)} in-range rows (total: {len(all_rows)})")

        if stop:
            print(f"  Reached records older than {CUTOFF}, stopping")
            break

    return all_rows


# ══════════════════════════════════════════════════════════════════════════════
# TAR CATANZARO – DataTables JSON API
# ══════════════════════════════════════════════════════════════════════════════

def scrape_tar_catanzaro() -> list:
    """
    Uses the Liferay DataTables AJAX endpoint discovered from the portlet JS.
    Iterates pages via start/length.
    """
    BASE = "https://www.giustizia-amministrativa.it"
    PAGE = f"{BASE}/provvedimenti-tar-catanzaro"
    print(f"\n[TAR Catanzaro] from {DATE_3D}")

    try:
        # Step 1: GET to build session
        r0 = SESSION.get(PAGE, timeout=25)
        r0.raise_for_status()
        soup0 = BeautifulSoup(r0.text, "lxml")

        target_form = next((f for f in soup0.find_all("form") if "administrative-acts" in f.get("action","")), None)
        if not target_form:
            print("  [!] Target form not found"); return []

        form_action = target_form.get("action")
        p_auth_m = re.search(r"p_auth=([^&]+)", form_action)
        p_auth = p_auth_m.group(1) if p_auth_m else ""

        prefix = "_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"
        form_data = {inp.get("name",""): inp.get("value","") for inp in target_form.find_all(["input","select"]) if inp.get("name")}
        form_data[f"{prefix}publishDateFrom"] = DATE_3D
        form_data[f"{prefix}publishDateTo"]   = DATE_TODAY

        # Step 2: POST to get the portlet with our filter (builds the additionalInfo JS var)
        r1 = SESSION.post(form_action, data=form_data,
                          headers={"Referer": PAGE}, timeout=20)
        r1.raise_for_status()

        search_url_m   = re.search(r'var searchURL = "([^"]+)"', r1.text)
        add_info_m     = re.search(r"d\.additionalInfo = '(\{[^']+\})'", r1.text)
        if not search_url_m or not add_info_m:
            print("  [!] Could not find searchURL or additionalInfo in portlet JS"); return []

        search_url   = search_url_m.group(1)
        add_info_str = add_info_m.group(1)
        if p_auth:
            search_url += f"&p_auth={p_auth}"

        print(f"  searchURL (short): ...{search_url[-60:]}")
        print(f"  additionalInfo: {add_info_str[:100]}")

        # Step 3: Paginate through DataTables results
        all_rows = []
        PAGE_SIZE = 100
        draw_num  = 1

        while True:
            payload = {
                "draw":    draw_num,
                "start":   len(all_rows),
                "length":  PAGE_SIZE,
                "columns": [
                    {"data": "nrgFascicolo",     "searchable": True, "orderable": True},
                    {"data": "sezione",          "searchable": True, "orderable": True},
                    {"data": "parte",            "searchable": True, "orderable": True},
                    {"data": "tipoUdienza",      "searchable": True, "orderable": True},
                    {"data": "dataUdienza",      "searchable": True, "orderable": True},
                    {"data": "numProvvedimento", "searchable": True, "orderable": True},
                    {"data": "dataPubblicazione","searchable": True, "orderable": True},
                    {"data": "tipoProvvedimento","searchable": True, "orderable": True},
                    {"data": "relatore",         "searchable": True, "orderable": True},
                    {"data": "presidente",       "searchable": True, "orderable": True},
                    {"data": "esito",            "searchable": True, "orderable": True},
                ],
                "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
                "search": {"value": "", "regex": False},
                "additionalInfo": add_info_str,
            }

            r2 = SESSION.post(
                search_url,
                data=json.dumps(payload),
                headers={
                    "Content-Type":   "application/json",
                    "Accept":         "application/json",
                    "X-Requested-With": "XMLHttpRequest",
                    "Referer":        PAGE,
                },
                timeout=20,
            )
            r2.raise_for_status()

            if not r2.content:
                print(f"  Empty response at start={len(all_rows)}"); break

            try:
                data = r2.json()
            except Exception:
                print(f"  Invalid JSON: {r2.text[:100]}"); break

            records   = data.get("data", [])
            total     = data.get("recordsFiltered", 0)
            print(f"  Batch start={len(all_rows)}: {len(records)} records (total={total})")

            for rec in records:
                # "parte" is the party field – the one we filter on
                parte  = rec.get("parte", "") or ""
                data_p = rec.get("dataPubblicazione", "") or ""
                nrg    = rec.get("nrgFascicolo", "") or ""
                nprov  = rec.get("numProvvedimento", "") or ""
                tipo   = rec.get("tipoProvvedimento", "") or ""

                # Build a meaningful "oggetto" string: Parte + NRG + tipo
                oggetto = parte
                if nrg:
                    oggetto = f"{parte} – NRG {nrg}"
                if tipo:
                    oggetto += f" ({tipo})"

                # Build link (detail page doesn't have a direct URL from the API,
                # use search result page as fallback)
                link = f"{PAGE}"

                all_rows.append({
                    "data":    data_p,
                    "oggetto": oggetto,
                    "link":    link,
                    "_parte":  parte,
                    "_nrg":    nrg,
                })

            if len(all_rows) >= total:
                break
            draw_num += 1
            if draw_num > 50:
                break

        return all_rows

    except Exception as e:
        print(f"  [TAR Error] {type(e).__name__}: {e}")
        import traceback; traceback.print_exc()
        return []


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 72)
    print(f"MONITORAGGIO ATTI CALABRIA – {TODAY.strftime('%d/%m/%Y')}")
    print(f"Periodo ASP / Regione  : {DATE_2D} → {DATE_TODAY}  (ultimi 2 giorni)")
    print(f"Periodo TAR            : {DATE_3D} → {DATE_TODAY}  (ultimi 3 giorni)")
    print("=" * 72)

    sheet_data  = {}
    sheet_notes = {}

    # ── ASP portals ──────────────────────────────────────────────────────────
    for portal_name, base_url in ASP_PORTALS:
        raw, err = scrape_asp(portal_name, base_url)
        filtered = [r for r in raw if matches(r["oggetto"])]
        print(f"  → {portal_name}: raw={len(raw)}, match={len(filtered)}")
        sheet_data[portal_name]  = filtered
        sheet_notes[portal_name] = err

    # ── Regione Calabria ─────────────────────────────────────────────────────
    raw_rc = scrape_regione_calabria()
    filtered_rc = [r for r in raw_rc if matches(r["oggetto"])]
    print(f"\n[Regione Calabria] raw={len(raw_rc)}, match={len(filtered_rc)}")
    sheet_data["Regione Calabria"]  = filtered_rc
    sheet_notes["Regione Calabria"] = ""

    # Show matching items
    for r in filtered_rc:
        print(f"  ✓ {r['data']} | {r['oggetto'][:70]}")

    # ── TAR Catanzaro ────────────────────────────────────────────────────────
    raw_tar = scrape_tar_catanzaro()
    # For TAR, filter on the "parte" field AND the whole oggetto string
    filtered_tar = [r for r in raw_tar if matches(r.get("_parte", r["oggetto"]))]
    print(f"\n[TAR Catanzaro] raw={len(raw_tar)}, match={len(filtered_tar)}")
    sheet_data["TAR Catanzaro"]  = filtered_tar
    sheet_notes["TAR Catanzaro"] = ""

    for r in filtered_tar:
        print(f"  ✓ {r['data']} | {r['oggetto'][:70]}")

    # ── Generate Excel ────────────────────────────────────────────────────────
    print(f"\nGenerating {OUT_FILENAME} …")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ORDER = [
        "ASP Cosenza",
        "ASP Catanzaro",
        "ASP Crotone",
        "ASP Reggio Calabria",
        "ASP Vibo Valentia",
        "Regione Calabria",
        "TAR Catanzaro",
    ]
    for name in ORDER:
        note = sheet_notes.get(name, "")
        rows = sheet_data.get(name, [])
        create_sheet(wb, name, rows, note=note)
        status = "INACCESSIBILE" if note else f"{len(rows)} atti"
        print(f"  {name:<26} → {status}")

    wb.save(OUT_PATH)
    print(f"\n✓ File salvato: {OUT_PATH}")

    total = sum(len(v) for v in sheet_data.values())
    accessible = sum(1 for n in sheet_notes.values() if not n)
    print(f"\nRiepilogo: {total} atti totali trovati su {accessible}/7 portali accessibili")


if __name__ == "__main__":
    main()
