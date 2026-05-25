# RUNBOOK — Pubblicazione LinkedIn giornaliera (Amodeo + S.Co Solution Consulting)

Questo file è la procedura che la sessione schedulata di Claude Code esegue ogni giorno.
Prompt del trigger: "Esegui automation/RUNBOOK.md per la data di oggi."

## 0. Setup di esecuzione

- Leggi `automation/cadence.json`: fuso, account, slot, numero di post per brand/giorno.
- Determina il giorno della settimana (Europe/Rome) e quanti post per brand vanno pubblicati oggi.
- Chiave Publer da env `PUBLER_API_KEY`. Browser disponibile via Playwright (`PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers`).

## 1. Quanti contenuti generare oggi

Per ciascun brand, da `weekly_posts`:
- valore 3 -> 2 post informativi + 1 long-form, agli slot `10:00/13:00/16:00`
- valore 1 -> 1 solo contenuto (il più forte e verificato del giorno), slot `20:00`

## 2. Scouting (fonti)

Ricerca web + fetch diretto delle fonti primarie. Priorità ultime 24-72 ore (allarga se il giorno è povero).
- Cyber/privacy/AI: acn.gov.it, garanteprivacy.it, csirt.gov.it, enisa.europa.eu, agid.gov.it; cybersecurity360.it, agendadigitale.eu, redhotcyber.com; eur-lex.europa.eu, normattiva.it.
- Sanità/pharma/MD/food: salute.gov.it, aifa.gov.it, ema.europa.eu, health.ec.europa.eu/medical-devices, agenas.gov.it, iss.it; efsa.europa.eu, food.ec.europa.eu, ilfattoalimentare.it; quotidianosanita.it, aboutpharma.com.
- HSE/ambiente/energia/ESG/governance/standard: inail.it, ispettorato.gov.it, mase.gov.it, enea.it, gse.it; accredia.it, uni.com; puntosicuro.it, esgnews.it, altalex.com.

Registra per ogni notizia: titolo, fonte, URL, data, sintesi, tema, (per SCO) settore.

## 3. Selezione

- SCO: copri almeno 2 macro-aree diverse; max 3 sullo stesso macro-tema; tieni traccia del mix settimanale.
- Amodeo: focus cyber/privacy/AI/ISO/NIS2/compliance/qualità sanitaria.
- Giorni da 1 post: scegli l'item singolo più rilevante e con la fonte più solida (priorità a scadenze imminenti).

## 4. GUARDRAIL DI VERIFICA (vincolante — "non dire stupidaggini")

- Ogni dato fattuale e ogni riferimento normativo (articolo, comma, ISO, Reg./Dir. UE, D.Lgs.) va confermato su fonte primaria nello stesso run.
- Se un riferimento non è confermabile: NON pubblicarlo come certo. O lo riformuli con riserva esplicita nel testo (vedi caso AI Act provvisorio), o scarti l'item.
- Una notizia da accordo/proposta non ancora adottata va sempre marcata come provvisoria nel testo.
- Nel report del giorno elenca per ogni affermazione la fonte URL. Nessuna pubblicazione di claim non verificati.

## 5. Scrittura (stile)

- Voce: Amodeo = consulente in prima persona/analitico; SCO = plurale ("noi"), claim implicito "Dare to change", mai prima persona singolare, mai auto-promozione.
- Hook che ferma lo scroll, poi riga vuota. Corpo 150-250 parole, frasi brevi. CTA = domanda aperta. 3-5 hashtag.
- Long-form: titolo <80 caratteri; struttura contesto/analisi/implicazioni/conclusione; **<= 2800 caratteri raw** (verifica con `finalize_md.mjs`).
- DIVIETI: niente emoji, niente clickbait, niente "buongiorno LinkedIn", niente auto-promo, niente frasi fatte, niente inglese forzato, niente riferimenti inventati.

## 6. Formattazione + disaiizzazione

- Scrivi il .md con marcatori `**...**` per il grassetto Unicode; converti e conta con `node tools/finalize_md.mjs <file>`.
- Rimuovi i tell AI: em dash, "questo è il cuore di", parallelismi "non solo X ma anche Y", "Chi X… Chi Y…" in chiusura, attribuzioni vaghe, conclusioni generiche.
- Whitelist (non modificare): riferimenti normativi, numeri/date/percentuali, hashtag, Unicode bold, nome S.Co Solution Consulting, claim "Dare to change".
- Salva il report sidecar `_report-modifiche.md`.

## 7. Card

`node tools/render_card.mjs <spec.json>` (viewport 1080x1080). Colori badge per tema (vedi prompt brand). SCO: wordmark + "Dare to change".

## 8. Pubblicazione (Publer)

Costruisci un manifest JSON: lista di `{account_id, text, media_path, scheduled_at}` con gli orari del giorno (ISO 8601 con offset, es. `2026-05-26T10:00:00+02:00`).
Esegui: `PUBLER_API_KEY=$KEY node tools/publish_publer.mjs manifest.json`.
Lo script carica i media, crea i post programmati e **riconcilia** la coda (il job può riportare failure transitorie: la verità è `GET /posts?state=scheduled`). Conferma che il numero di post in coda corrisponda all'atteso.

## 9. Output

Salva tutto in `YYYY-MM-DD/`, committa e pusha sul branch di lavoro. Niente chiavi nei file versionati.
