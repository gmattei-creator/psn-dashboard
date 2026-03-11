---
name: riepilogo-email-pmo-psn
description: Dashboard PSN – Monitor Ferrari/Sogei (ore 9 e 19, lun-ven)
---

In italiano. Esegui il riepilogo giornaliero del progetto PSN (Polo Strategico Nazionale) per DPC.

═══════════════════════════════════════════
STEP 0 — CARICA STATO E CONFIGURAZIONE
═══════════════════════════════════════════

Leggi /mnt/PSN/last_run_state.json e /mnt/PSN/github_config_psn.txt.
Estrai:
  - processed_message_ids (lista ID già elaborati)
  - processed_sent_ids
  - gantt_items, log_entries, run_history, verticali_stato
  - last_run_date, last_run_time
  - GITHUB_TOKEN, GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH

Calcola last_run_epoch in Unix timestamp (secondi), timezone Europe/Rome:
  Usa Python:
    from datetime import datetime
    import pytz
    tz = pytz.timezone("Europe/Rome")
    dt = tz.localize(datetime.strptime(f"{last_run_date} {last_run_time}", "%Y-%m-%d %H:%M"))
    last_run_epoch = int(dt.timestamp())
  Se pytz non disponibile: usa offset +1 (CET) o +2 (CEST, da ultima domenica di marzo).

Assicura dipendenze Python installate:
  pip install reportlab openpyxl --break-system-packages -q
  (esegui solo se non già disponibili; ignora output)


═══════════════════════════════════════════
STEP 1 — RACCOLTA DATI IN PARALLELO
═══════════════════════════════════════════

Esegui i due blocchi seguenti IN PARALLELO (stessa chiamata):

── BLOCCO A: EMAIL NUOVE ─────────────────
Usa Gmail API con query:
  (from:eferrari@sogei.it OR from:sogei.it) after:{last_run_epoch} -in:sent
Filtra: tieni solo messaggi il cui ID NON è in processed_message_ids.
Risultato → nuove_email_ids (lista di ID)

── BLOCCO B: CALENDARIO DPC-LAVORO ───────
Recupera eventi oggi dal calendario:
  ID: pu3ln0ksa737gen5h50ijgrglcg80br5@import.calendar.google.com
  Intervallo: 00:00–23:59 oggi (Europe/Rome)
Risultato → cal_oggi (lista eventi con id, summary, start, end, description)

Attendi completamento di entrambi i blocchi.


═══════════════════════════════════════════
STEP 1.5 — LETTURA CORPO COMPLETO EMAIL NUOVE
═══════════════════════════════════════════

SOLO SE nuove_email_ids non è vuoto:
  Per ogni ID in nuove_email_ids, chiama gmail_read_message(messageId=id).
  Costruisci nuove_email = lista di oggetti:
    { id, oggetto, data, mittente, corpo_testo, snippet, allegati: [{nome, attachmentId}] }
  Nota: il corpo completo permette estrazione attività più accurata al STEP 3.

SOLO SE nuove_email_ids non è vuoto — BLOCCO C: ALLEGATI EXCEL FERRARI ─
  Query Gmail: from:eferrari@sogei.it has:attachment (filename:xlsx OR filename:xls OR filename:xlsm)
  Per ogni email con allegato Excel il cui ID NON è in processed_message_ids:
    a) Scarica via Gmail API: GET .../messages/{messageId}/attachments/{attachmentId}
       campo "data" = base64url → decodifica → salva /tmp/allegato_{id}.xlsx
    b) Importa openpyxl; parsa workbook; mappa colonne Gantt (vedi sezione GANTT sotto)
  Risultato → nuovi_gantt_excel (lista voci Gantt estratte da Excel)
  Se nessun nuovo Excel trovato: nuovi_gantt_excel = []


═══════════════════════════════════════════
STEP 2 — DECISIONE: RUN PIENA O LEGGERA
═══════════════════════════════════════════

Calcola un "fingerprint" dello stato corrente:
  import hashlib, json
  fp_input = json.dumps({
    "ids":    sorted(list(processed_message_ids) + [e["id"] for e in nuove_email]),
    "cal":    sorted([ev["id"] for ev in cal_oggi]),
    "gantt":  sorted([g["nome"] for g in gantt_items + nuovi_gantt_excel])
  }, sort_keys=True)
  fingerprint = hashlib.md5(fp_input.encode()).hexdigest()

Confronta con il fingerprint salvato in state.json (campo "last_fingerprint").

SE fingerprint == last_fingerprint (nulla di nuovo):
  → RUN LEGGERA:
    1. Aggiungi UNA voce "system" al log (in cima):
       { "tipo":"system", "data":oggi, "ora":ora_attuale,
         "titolo":"Run automatica – {data} {ora} (nessuna novità)",
         "descrizione":"Fingerprint invariato. Nessuna nuova email, Gantt o evento calendario.",
         "badge":"🖥️ Sistema" }
    2. Aggiorna last_run_date e last_run_time nello state.json
    3. Salva /mnt/PSN/last_run_state.json LOCALMENTE (solo scrittura file)
    4. NON generare PDF, NON aggiornare HTML, NON fare push GitHub, NON creare bozza email
    5. FINE — termina qui.

SE fingerprint != last_fingerprint (ci sono novità):
  → RUN PIENA: prosegui con gli step 3-7.


═══════════════════════════════════════════
STEP 3 — ANALISI EMAIL E AGGIORNAMENTO GANTT
═══════════════════════════════════════════

Classifica ogni nuova email per priorità (controlla oggetto E corpo completo):
  🔴 CRITICO:    "bloccante","KO","fermo","urgente","escalation","scadenza oggi","entro oggi","fermo tutto"
  🟡 ATTENZIONE: "entro","scadenza","sollecito","in attesa","pendente","verifica","VLAN",
                  "configurazione","in corso","upgrade","manleva","manvela","firma","approvazione"
  🟢 INFO:       tutto il resto

Estrai attività implicite dal corpo completo di ogni email nuova Ferrari:
  - "entro [data]" / "scadenza [data]" → Fine = data estratta, stato = pending
  - "in attesa di [X]" → stato = pending, owner = DPC
  - "bloccato/fermo su [X]" → stato = blocked
  - "completato/chiuso/risolto [X]" → stato = completed
  - "dobbiamo/bisogna/serve [fare X]" → nuova attività pending
  - Owner default: "Ferrari/Sogei"
  - Verticale: inferisci da keyword (AWS/DB/VPN/IAM/Sicurezza/Finanziamento/SAL/PSN/Fornitori)

Merge Gantt (Excel ha priorità su testo, fuzzy match >70%):
  gantt_aggiornato = merge(gantt_items, nuovi_gantt_excel, attività_da_testo)

Sync verticali_stato: per ogni Gantt item con stato "completed",
  decrementa aperti e incrementa chiusi nel verticale corrispondente (se aperti > 0).
  Aggiungi il verticale a verticali_stato se non esiste.

Aggiungi ID nuove email a processed_message_ids.


═══════════════════════════════════════════
ANALISI E CLASSIFICAZIONE EMAIL
═══════════════════════════════════════════

Label email: "Email PMO" (mai "Email RTI")

Colonne Gantt da Excel Ferrari (best-effort mapping):
  Attività/Task/Descrizione → nome | Stato/Status → stato | Inizio/Start → data inizio
  Fine/Scadenza/End → data fine | Owner/Responsabile → owner | %/Avanzamento → pct
  Verticale/Area/Categoria → verticale

Normalizzazione stato:
  completato/chiuso/done/risolto → completed | in corso/wip/attivo → in-progress
  bloccato/ko/fermo → blocked | pendente/todo/da fare → pending

Badge fonte: "📎 Excel" se da allegato, "✉️ Email" se da testo email.


═══════════════════════════════════════════
STEP 4 — GENERA PDF REPORT (solo run piena)
═══════════════════════════════════════════

Usa reportlab (già installato al STEP 0). Output: /mnt/PSN/PSN_Report_{YYYYMMDD}_{HHMM}.pdf

Struttura:
  - Header gradiente + "DIFFUSIONE LIMITATA"
  - KPI row (email totali, gantt aperti, nuove email oggi, riunioni oggi)
  - Riunioni oggi (da cal_oggi)
  - Email PMO (ultime 10-15, con badge priorità; nuove email evidenziate in giallo)
  - Gantt attività aperte (nuove voci evidenziate)
  - Stato verticali
  - Azioni prioritarie
  - Footer "DIFFUSIONE LIMITATA"


═══════════════════════════════════════════
STEP 5 — AGGIORNA DASHBOARD HTML (solo run piena)
═══════════════════════════════════════════

Singolo file index.html con 7 TAB (JS puro, nessun reload):

── LOGIN OVERLAY ─────────────────────────
SOLO PASSWORD: DashDpc@26 (nessun username)
Errore: "❌ Password non valida." | Supporto tasto Invio

── 7 TAB ─────────────────────────────────
  1. 📊 Overview     – KPI + azioni prioritarie + ultime 4 email
  2. 📧 Email PMO    – Email Ferrari/Sogei (label "Email PMO", MAI "Email RTI")
  3. 📅 Calendario   – Riunioni del giorno
  4. 📈 Gantt        – Da Excel Ferrari + testo email; badge "📎 Excel" o "✉️ Email" per riga
  5. 🏗 Verticali    – Stato per verticale con barre progresso
  6. 🕐 Storico      – Storico run automatici (max 20)
  7. 📋 Log          – Timeline cronologica di tutte le attività (vedi sotto)

── TAB LOG ───────────────────────────────
Timeline cronologica raggruppata per giorno. Tipi: email | cal | action | excel | gantt | system
Toolbar con filtri JS (Tutti / Email / Riunioni / Azioni / Excel / Gantt / Sistema) + ricerca full-text.
Nuovi eventi SEMPRE in cima; storici conservati.

── STILE ─────────────────────────────────
Palette: #1a3a5c / #2563eb / #f1f5f9. Header gradiente + badge LIVE.
Tab attivo: bordo inferiore blu + sfondo #eff6ff. Responsive. Footer: Diffusione Limitata.
Nuove voci (email/Gantt) evidenziate con bordo/sfondo giallo (#fef3c7) e label 🆕.


═══════════════════════════════════════════
STEP 6 — AGGIORNA STATE.JSON E PUSH GITHUB (COMMIT UNIFICATO)
═══════════════════════════════════════════

Aggiorna state.json con (applica i cap PRIMA di salvare):
  processed_message_ids  → mantieni max 200 (rimuovi i più vecchi se > 200)
  log_entries            → mantieni max 50  (rimuovi i più vecchi se > 50)
  run_history            → mantieni max 20  (rimuovi i più vecchi se > 20)
  processed_sent_ids, gantt_items, verticali_stato,
  last_run_date, last_run_time, last_fingerprint, contatori vari.

Salva /mnt/PSN/last_run_state.json localmente.

Push su GitHub — USA UN SINGOLO COMMIT via Git Trees API (transazione atomica):

  # 1. Ottieni SHA dell'ultimo commit su main
  GET /repos/{user}/{repo}/git/refs/heads/{branch}
  → estrai commit_sha

  # 2. Ottieni base tree SHA dal commit
  GET /repos/{user}/{repo}/git/commits/{commit_sha}
  → estrai base_tree_sha

  # CIRCUIT BREAKER per index.html:
  #   Calcola git blob SHA locale (formato git):
  #     sha1("blob " + str(len(html_bytes)) + "\x00" + html_bytes)
  #   Confronta con il campo "sha" di:
  #     GET /repos/{user}/{repo}/contents/index.html?ref={branch}
  #   Se coincidono → html_unchanged = True (ometti index.html dal tree)
  #   Se diversi   → html_unchanged = False (includi index.html nel tree)

  # 3. Crea un nuovo tree
  POST /repos/{user}/{repo}/git/trees
  body: {
    "base_tree": base_tree_sha,
    "tree": [
      # Sempre incluso:
      {"path": "data/state.json", "mode": "100644", "type": "blob", "content": <json_string>},
      # Solo se html_unchanged == False:
      {"path": "index.html",      "mode": "100644", "type": "blob", "content": <html_string>}
    ]
  }
  → estrai new_tree_sha

  # 4. Crea il commit
  POST /repos/{user}/{repo}/git/commits
  body: {"message": "chore: run {YYYYMMDD} {HH:MM}", "tree": new_tree_sha, "parents": [commit_sha]}
  → estrai new_commit_sha

  # 5. Aggiorna il ref (fast-forward)
  PATCH /repos/{user}/{repo}/git/refs/heads/{branch}
  body: {"sha": new_commit_sha}


═══════════════════════════════════════════
STEP 7 — BOZZA EMAIL RIEPILOGO (solo run piena)
═══════════════════════════════════════════

Crea bozza Gmail a Giordano.Mattei@protezionecivile.it con:
  Oggetto: "📊 Riepilogo PSN – {DD/MM/YYYY} ore {HH:MM} | {N} nuove email | {M} riunioni oggi"
  Corpo HTML: riunioni del giorno, azioni prioritarie, Gantt aperti (tabella),
              link dashboard (https://gmattei-creator.github.io/psn-dashboard/),
              nome PDF generato.
  Footer: "🔒 DIFFUSIONE LIMITATA"


═══════════════════════════════════════════
REGOLE IMPORTANTI
═══════════════════════════════════════════

- NON riprocessare email già in processed_message_ids
- BLOCCO C (Excel) e STEP 1.5 (lettura corpo) girano SOLO se nuove_email_ids non è vuoto
- Run leggera: SOLO log system + save locale — NESSUN push GitHub, NESSUN PDF, NESSUN HTML
- Commit GitHub sempre unificato (Git Trees API) — mai due PUT separati
- Label email: "Email PMO" (mai "Email RTI")
- Login: SOLO password "DashDpc@26", NESSUN username
- Gantt: Excel Ferrari (priorità su testo); badge fonte per ogni riga
- Log: nuovi eventi SEMPRE in cima; cap a 50 voci prima di salvare
- processed_message_ids: cap a 200 voci (rimuovi le più vecchie)
- Classificazione "Diffusione Limitata" su tutti gli output
- GitHub config in /mnt/PSN/github_config_psn.txt
- Nuove voci Gantt/email nella dashboard evidenziate con 🆕 e bordo giallo
- Verticali: sincronizza aperti/chiusi quando un Gantt item cambia stato
