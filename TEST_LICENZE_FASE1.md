# Test licenze — Fase 1 (Editor Tecnico)

Verifica dell'enforcement licenze introdotto in `main.py` (commit `c1094c3`) + `server_time` lato server (`licenze_verifica.php`, commit sito `db8dd4d`).

## Cosa è stato implementato (in breve)
- All'avvio l'app chiama il server: **il server comanda**. Se non c'è rete, usa una **cache firmata** salvata nel `settings.ini`.
- Offline si lavora solo se: firma integra **AND** entro la finestra offline (default **48h**) **AND** non oltre la scadenza **AND** orologio non spostato indietro.
- Blocco **morbido**: l'app si apre, ma il workspace e le funzioni di lavoro restano disattivati, con messaggio.

## Dove sta il settings.ini
- Windows: di norma `C:\Users\<utente>\AppData\Roaming\<...>\settings.ini` (cartella `AppDataLocation`),
  oppure fallback `C:\Users\<utente>\.EditorTecnico\settings.ini`.
- Chiavi rilevanti: `license/key`, `license/server_url`, `license/good_scadenza`, `license/good_ora_server`,
  `license/good_ora_locale`, `license/good_firma`, `license/max_offline_ore` (opzionale, default 48).

> Suggerimento: prima di toccare il `.ini` a mano, fanne una copia.

---

## TEST 1 — Online, licenza valida (caso base)
1. Connessione internet attiva, chiave valida configurata.
2. Avvia: `python main.py`.
3. **Atteso:** nessun popup di blocco; workspace selezionabile/abilitato; barra di stato mostra "Workspace attivo".
4. Apri il `settings.ini`: devono essere presenti e valorizzate `license/good_scadenza`, `license/good_ora_server`, `license/good_ora_locale`, `license/good_firma`.

## TEST 2 — Offline entro la finestra (48h)
1. Esegui prima il TEST 1 (per seminare la cache firmata).
2. Stacca la rete (disattiva Wi-Fi/ethernet).
3. Riavvia l'app.
4. **Atteso:** l'app resta **abilitata** (siamo entro 48h dall'ultima verifica online e prima della scadenza).

## TEST 3 — Manomissione del settings.ini (firma HMAC)
1. Con app chiusa, apri il `settings.ini` e modifica a mano un valore della cache, es. cambia `license/good_scadenza` in una data più avanti.
2. Riavvia l'app **offline**.
3. **Atteso:** **blocco** con messaggio "Dati di licenza alterati. Collegati a internet per riverificare la licenza."

## TEST 4 — Orologio spostato indietro (anti-rollback)
1. Esegui il TEST 1 (cache seminata), poi chiudi l'app e stacca la rete.
2. Sposta la data/ora di sistema **nel passato** (prima dell'ultima verifica).
3. Riavvia l'app offline.
4. **Atteso:** **blocco** con messaggio "Orologio di sistema modificato...".
5. Ripristina data/ora corrette al termine.

## TEST 5 — Oltre la finestra offline (>48h)
1. Cache seminata (TEST 1), app chiusa, rete staccata.
2. Sposta la data/ora **avanti di oltre 48h** (ma prima della scadenza licenza).
3. Riavvia offline.
4. **Atteso:** **blocco** con "Limite di utilizzo offline superato (48 ore)...".
5. Ripristina data/ora.

## TEST 6 — Licenza scaduta
1. Con una licenza la cui scadenza è vicina, oppure spostando la data di sistema **oltre la scadenza**.
2. Riavvia (online oppure offline).
3. **Atteso:** **blocco** con "Licenza scaduta." (online: lo dice il server; offline: lo calcola la cache).
4. Ripristina data/ora.

## TEST 7 — Licenza revocata/sospesa lato CMS (online)
1. Dal CMS metti la licenza in stato `revocata` o `sospesa`.
2. Con rete attiva, riavvia l'app.
3. **Atteso:** **blocco** immediato con il messaggio del server (non aspetta la scadenza).
4. Riporta la licenza ad `attiva` per i test successivi.

## TEST 8 — Verifica manuale dal menu
1. Dalla finestra "Verifica licenza" premi **Verifica licenza** con rete attiva e chiave valida.
2. **Atteso:** esito ok, workspace si riabilita, la cache `license/good_*` viene rigenerata.

## TEST 9 — Primo avvio dopo l'aggiornamento, da offline (comportamento atteso)
1. Su una macchina che non ha mai eseguito questa versione (nessuna chiave `license/good_*`), avvia **senza rete**.
2. **Atteso:** **blocco** con "Nessuna verifica online valida disponibile...". È corretto: serve **una** verifica online iniziale per seminare la cache.

---

## Note
- La finestra offline è configurabile via chiave `license/max_offline_ore` nel `settings.ini` (default 48). In Fase 2 sarà pilotata dal CMS.
- Se il server non restituisce `server_time` (versione vecchia del PHP), il desktop ripiega sull'ora locale al momento del sync: tutto funziona, ma l'anti-rollback è un filo più debole. Il PHP aggiornato è già in produzione.
