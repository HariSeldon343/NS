# SETUP — Attivazione dell'automazione giornaliera

Stato: il codice è pronto. Mancano azioni che solo il titolare può fare. In ordine.

## 1. Sbloccare la scrittura sul repo (obbligatorio)

Oggi `git push` e l'API GitHub rispondono 403 (sola lettura). Senza scrittura non si può
depositare il codice né farlo girare in automatico.
- Concedi alla GitHub App di Claude Code il permesso **Contents: write** su `hariseldon343/ns`
  (GitHub → Settings → Integrations/Applications → Claude Code → repository access), oppure
- ricrea l'ambiente Claude Code on the web autorizzando lo scope di scrittura sul repo.

## 2. Inserire la chiave Publer come secret (non in chiaro)

Nelle impostazioni dell'ambiente Claude Code on the web, aggiungi una variabile/secret:
- `PUBLER_API_KEY` = (la tua API key Publer Business)
- opzionale `PUBLER_WORKSPACE_ID` = `6a13de8cadf35ab5536f6cf6`

Importante: **rigenera la chiave Publer** (è stata condivisa in chat) e usa quella nuova qui.

## 3. Creare il trigger giornaliero

Nell'interfaccia web di Claude Code, crea una sessione **schedulata** ricorrente:
- frequenza: ogni giorno
- orario consigliato: ~07:30 Europe/Rome (lascia margine prima dello slot delle 10:00)
- prompt della sessione: `Esegui automation/RUNBOOK.md per la data di oggi.`

La sessione farà scouting verificato, scrittura, card, e programmerà i post su Publer agli orari
della cadenza. Publer pubblica da solo su profilo e pagina.

## 4. Cadenza (riepilogo)

Vedi `automation/cadence.json`.
- Amodeo: 3 post (10/13/16) Mar e Gio; 1 post (20:00) Lun/Mer/Ven/Sab/Dom. ~11/sett.
- SCO: 3 post (10/13/16) Lun/Mer/Ven/Sab/Dom; 1 post (20:00) Mar e Gio. ~17/sett.

## 5. Verifica del primo giro automatico

Il primo giorno, controlla in Publer che i post programmati corrispondano alla cadenza e che le
card siano allegate. Da lì non devi più intervenire.

## Note

- Guardrail anti-errori attivo: nessun claim non verificato viene pubblicato come certo
  (le notizie provvisorie escono con riserva esplicita nel testo). Vedi RUNBOOK sezione 4.
- I file `tools/render_card.mjs`, `tools/finalize_md.mjs`, `tools/publish_publer.mjs` sono già pronti e collaudati.
