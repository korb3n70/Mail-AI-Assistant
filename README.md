# Copezzot Mail AI Assistant

Client email leggero per Windows (Outlook COM) con integrazione Claude AI
per composizione e risposta alle email. PowerShell 5 / WinForms, nessuna
dipendenza esterna oltre Outlook desktop.

## Requisiti

- Windows con Outlook desktop installato e configurato (Office 365, Gmail,
  IMAP o altro — lo script non gestisce mai le credenziali direttamente,
  usa la sessione COM già autenticata di Outlook)
- PowerShell 5.1 o superiore
- Una API key Anthropic ([console.anthropic.com](https://console.anthropic.com))

## File del progetto

| File | Descrizione |
|---|---|
| `clientmail_outlook_v20.ps1` | Script principale — tutta l'applicazione |
| `favicon.ico` | Icona per lo shortcut Windows |
| `copezzot_logo.jpeg` | Logo mostrato nell'header dell'app |
| `changelog.txt` | Storico versioni, letto e mostrato in-app |
| `esegui.cmd` | Batch di lancio (evita problemi di ExecutionPolicy) |
| `Copezzot.lnk` | Shortcut pronto: punta a `esegui.cmd`, icona `favicon.ico` |

**Tutto il progetto gira da `%USERPROFILE%\MailClient\`.** Lo script
trova la propria cartella dinamicamente (`$MyInvocation.MyCommand.Path`),
quindi funziona ovunque venga copiato — ma `favicon.ico`,
`copezzot_logo.jpeg` e `changelog.txt` devono restare **nella stessa
cartella** dello script `.ps1`, qualunque essa sia.

## Avvio

Copia l'intero contenuto di questa cartella in `%USERPROFILE%\MailClient\`,
poi:

- Doppio click su `Copezzot.lnk` (consigliato — nessun terminale visibile,
  icona personalizzata), oppure
- Doppio click su `esegui.cmd`, oppure
- Da terminale:
  ```powershell
  powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\MailClient\clientmail_outlook_v20.ps1"
  ```

Il file `.lnk` incluso è già configurato con i percorsi corretti per
questo setup — se sposti il progetto altrove, ricrea lo shortcut.

## Dati generati a runtime (non versionati)

Al primo avvio lo script crea `%USERPROFILE%\MailClient\` con:

- `creds.xml` — API key cifrata con DPAPI (legata all'account Windows,
  illeggibile su un altro PC/utente)
- `app_settings.json` — preset AI, colori, MOTD, font
- `contacts_cache.json` — cache indirizzi per l'autocomplete
- `drafts.json` — bozze locali
- `style_<email>.json` — profilo di stile di scrittura personale (opzionale)

Questi file **non vanno mai committati** — vedi `.gitignore`. Il tab
Impostazioni ha un pulsante "Azzera dati personali" per cancellarli tutti
prima di condividere lo script con qualcun altro.

## Architettura in breve

- **Outlook COM automation** (`New-Object -ComObject Outlook.Application`)
  per leggere/inviare email — funziona con qualsiasi account configurato
  in Outlook, incluso multi-account (selezionabile in Impostazioni)
- **Claude API** via `HttpWebRequest` (non `WebClient`, per evitare
  problemi di encoding UTF-8/BOM tipici di PowerShell 5)
- **WebBrowser control** (IE engine) come editor HTML rich-text per
  Compose e Reply, con toolbar di formattazione custom
- **Apprendimento stile personale**: confronta bozza AI vs versione
  modificata dall'utente ad ogni invio, aggiorna un profilo testuale
  usato nelle generazioni successive (opt-in, disattivabile)

## Nota su PowerShell 5 e closure

Il codice contiene diversi pattern `$script:xxx = $control` subito dopo
la creazione di controlli WinForms. Non è ridondanza: in PowerShell 5 gli
scriptblock degli `Add_Click`/`Add_TextChanged` **non catturano** le
variabili locali della funzione in cui sono definiti — solo lo scope
`$script:` è visibile in modo affidabile. Se aggiungi nuovi controlli con
handler, segui lo stesso pattern.

## Licenza

Uso interno.
