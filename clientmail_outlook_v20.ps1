# ============================================================
# Light Mail Client + AI - Versione Outlook COM v20
# v20: pannello credito API, selezione modello AI,
#      chiarimento Admin Key per statistiche utilizzo
# ============================================================

# Caricamento assembly compatibile con PowerShell 5.x e 7+
$assemblies = @("System.Windows.Forms", "System.Drawing", "System.Security")
foreach ($asm in $assemblies) {
    try { Add-Type -AssemblyName $asm -ErrorAction SilentlyContinue } catch {}
}

# ============================================================
# PERCORSI FILE
# ============================================================
$rootPath     = "$env:USERPROFILE\MailClient"
$draftsFile   = "$rootPath\drafts.json"
$credsFile    = "$rootPath\creds.xml"
$contactsFile = "$rootPath\contacts_cache.json"
$settingsFile = "$rootPath\app_settings.json"

if (-not (Test-Path $rootPath)) {
    New-Item -ItemType Directory -Force -Path $rootPath | Out-Null
}

# -- Cache contatti --------------------------------------------
function Load-Contacts {
    $map = @{}
    if (-not (Test-Path $contactsFile)) { return $map }
    try {
        $raw = [System.IO.File]::ReadAllText($contactsFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($p in $raw.PSObject.Properties) {
            $map[$p.Name] = $p.Value
        }
    } catch {}
    return $map
}

function Save-Contacts {
    param($map)
    try {
        $obj = [PSCustomObject]@{}
        foreach ($k in ($map.Keys | Sort-Object)) {
            $obj | Add-Member -NotePropertyName $k -NotePropertyValue $map[$k] -Force
        }
        $json = $obj | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($contactsFile, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

function Add-ContactToCache {
    param($email, $name)
    if (-not $email -or $email -notmatch "@") { return }
    $email = $email.Trim().ToLower()
    $name  = if ($name -and $name -ne $email) { $name.Trim() } else { $email }
    $map   = Load-Contacts
    if ($map[$email]) {
        $map[$email].count = [int]$map[$email].count + 1
        $map[$email].last  = (Get-Date).ToString("yyyy-MM-dd")
        if ($name -and $name -ne $email) { $map[$email].name = $name }
    } else {
        $map[$email] = [PSCustomObject]@{ name=$name; count=1; last=(Get-Date).ToString("yyyy-MM-dd") }
    }
    Save-Contacts $map
}

function Search-Contacts {
    param($query, $maxResults=8)
    $map = Load-Contacts
    $q   = $query.ToLower()
    $map.GetEnumerator() |
        Where-Object {
            $_.Key -like "*$q*" -or
            $_.Value.name.ToLower() -like "*$q*"
        } |
        Sort-Object { -[int]$_.Value.count } |
        Select-Object -First $maxResults |
        ForEach-Object {
            [PSCustomObject]@{
                Email   = $_.Key
                Name    = $_.Value.name
                Count   = $_.Value.count
                Display = "$($_.Value.name) <$($_.Key)>"
            }
        }
}

# Inizializza cache in memoria all'avvio
$script:contactsCache = Load-Contacts

# -- Impostazioni app (preset AI + tema) ----------------------
function Load-AppSettings {
    $defaults = @{
        Model       = 0
        Language    = "Italiano"
        Tone        = "Conciso"
        HeaderColor = "255,255,255"
        FormBgColor = "255,255,255"
        AppTitle    = "Mail AI Assistant"
        FormTitle   = "Mail AI Assistant"
        MotdFont    = "Segoe UI"
        MotdSize    = 14
        MotdColor   = "30,30,30"
        SignatureName = ""
        StyleLearningEnabled = $false
        StyleUsageEnabled    = $true
        MaxEmails            = 30
        SelectedAccount      = ""
    }
    if (-not (Test-Path $settingsFile)) { return $defaults }
    try {
        $raw = [System.IO.File]::ReadAllText($settingsFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($k in @($defaults.Keys)) {
            try {
                $val = $raw.PSObject.Properties[$k]
                if ($val -ne $null -and $val.Value -ne $null) {
                    $defaults[$k] = $val.Value
                }
            } catch {}
        }
    } catch {}
    return $defaults
}

function Save-AppSettings {
    param($settings)
    try {
        $json = $settings | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($settingsFile, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

function Parse-Color {
    param([string]$csv)
    try {
        $p = $csv -split ","
        return [System.Drawing.Color]::FromArgb([int]$p[0],[int]$p[1],[int]$p[2])
    } catch { return [System.Drawing.Color]::White }
}

$script:appSettings = Load-AppSettings

# ============================================================
# STILE PERSONALE - apprendimento dal confronto AI vs modifiche utente
# ============================================================
function Get-UserEmailAddress {
    try {
        $ol = Get-OutlookInstance
        $ns = $ol.GetNamespace("MAPI")
        $cu = $ns.CurrentUser
        try {
            $exUser = $cu.AddressEntry.GetExchangeUser()
            if ($exUser -and $exUser.PrimarySmtpAddress) { return $exUser.PrimarySmtpAddress }
        } catch {}
        if ($cu.Address -match "@") { return $cu.Address }
    } catch {}
    return "utente_sconosciuto"
}

function Get-StyleFilePath {
    $email = Get-UserEmailAddress
    $safe  = ($email -replace '[^\w\.\-@]', '_')
    return "$rootPath\style_$safe.json"
}

function Load-StyleProfile {
    $path = Get-StyleFilePath
    $default = @{
        Email       = Get-UserEmailAddress
        Summary     = ""
        History     = @()
        LastUpdated = ""
    }
    if (-not (Test-Path $path)) { return $default }
    try {
        $raw = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        $default.Email       = if ($raw.Email) { $raw.Email } else { $default.Email }
        $default.Summary     = if ($raw.Summary) { $raw.Summary } else { "" }
        $default.LastUpdated = if ($raw.LastUpdated) { $raw.LastUpdated } else { "" }
        if ($raw.History) { $default.History = @($raw.History) }
    } catch {}
    return $default
}

function Save-StyleProfile {
    param($profile)
    try {
        $path = Get-StyleFilePath
        $json = $profile | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($path, $json, [System.Text.Encoding]::UTF8)
    } catch {}
}

# Chiede a Claude di confrontare la versione AI con quella modificata dall'utente
# e aggiornare un riassunto testuale dello stile personale
function Update-StyleProfile {
    param($aiText, $userText, $apiKey)

    if ([string]::IsNullOrWhiteSpace($apiKey)) { return }
    if ($aiText.Trim() -eq $userText.Trim()) { return }   # nessuna modifica, niente da imparare

    $profile = Load-StyleProfile
    $prevSummary = if ($profile.Summary) { $profile.Summary } else { "(nessuno stile registrato ancora)" }

    $sysMsg = "You analyze how a user edits AI-generated emails to learn their personal writing style. " +
              "You will be given: the previous style summary, the AI-generated draft, and the user's final edited version. " +
              "Update the style summary to reflect any new patterns you observe (tone, formality, sentence length, structure, punctuation habits, level of detail). " +
              "CRITICAL: Describe the STYLE ABSTRACTLY, never in the specific language of the sample. " +
              "For greetings and sign-offs, describe the REGISTER (e.g. 'informal greeting', 'warm but brief opening', 'formal closing with full name') " +
              "and NEVER quote the literal words used (do not write things like 'uses hi' or 'uses Best regards' or 'usa Ciao'), because the email may be written in a different language each time and literal foreign words would be wrongly copied verbatim. " +
              "Keep the summary concise (max 100 words), written in Italian, as a set of practical writing directives an assistant could follow regardless of the target language. " +
              "If the edits are minor or don't reveal a clear pattern, keep the summary mostly unchanged. " +
              "Output ONLY the updated summary text, nothing else."

    $usrMsg = "Previous style summary:`n$prevSummary`n`n" +
              "AI-generated draft:`n$aiText`n`n" +
              "User's final edited version:`n$userText`n`n" +
              "Updated style summary:"

    try {
        $payloadObj = @{
            model      = "claude-haiku-4-5-20251001"
            max_tokens = 300
            system     = $sysMsg
            messages   = @(@{ role = "user"; content = $usrMsg })
        }
        $payload   = $payloadObj | ConvertTo-Json -Depth 10
        $payload   = $payload -replace '\\u003c','<' -replace '\\u003e','>' -replace '\\u0027',"'" -replace '\\u0026','&'
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $bodyBytes = $utf8NoBom.GetBytes($payload)

        $req = [System.Net.HttpWebRequest]::Create("https://api.anthropic.com/v1/messages")
        $req.Method = "POST"; $req.ContentType = "application/json; charset=utf-8"
        $req.ContentLength = $bodyBytes.Length; $req.Timeout = 15000
        $req.Headers.Add("x-api-key", $apiKey)
        $req.Headers.Add("anthropic-version", "2023-06-01")
        $s = $req.GetRequestStream(); $s.Write($bodyBytes,0,$bodyBytes.Length); $s.Close()
        $resp   = $req.GetResponse()
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream(), $utf8NoBom)
        $json   = $reader.ReadToEnd(); $reader.Close(); $resp.Close()
        $result = $json | ConvertFrom-Json
        $newSummary = $result.content[0].text.Trim()

        $profile.Summary     = $newSummary
        $profile.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        $profile.History     = @($profile.History) + @([PSCustomObject]@{
            Date    = (Get-Date).ToString("yyyy-MM-dd HH:mm")
            Summary = $newSummary
        })
        # Mantieni solo le ultime 30 osservazioni per non far crescere troppo il file
        if ($profile.History.Count -gt 30) {
            $profile.History = $profile.History[-30..-1]
        }
        Save-StyleProfile $profile

        # Aggiorna la UI del tab stile se esiste
        try {
            if ($script:txtStyleSummaryRef) { $script:txtStyleSummaryRef.Text = $newSummary }
            if ($script:lstStyleHistoryRef) {
                $script:lstStyleHistoryRef.Items.Clear()
                foreach ($h in ($profile.History | Sort-Object Date -Descending)) {
                    $script:lstStyleHistoryRef.Items.Add("$($h.Date)  -  $($h.Summary)") | Out-Null
                }
            }
        } catch {}
    } catch {
        # Fallimento silenzioso - non deve mai bloccare l'invio dell'email
    }
}

# ============================================================
# CREDENZIALI CIFRATE CON DPAPI (solo API Key per AI)
# ============================================================
function Save-Credentials {
    param($apiKey)
    $encApiKey = ConvertFrom-SecureString (ConvertTo-SecureString $(
        if ($apiKey) { $apiKey } else { "none" }
    ) -AsPlainText -Force)
    [PSCustomObject]@{
        ApiKey = $encApiKey
    } | Export-Clixml -Path $credsFile
}

function Load-Credentials {
    if (-not (Test-Path $credsFile)) { return $null }
    try {
        $c = Import-Clixml -Path $credsFile
        $decApi = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(
                (ConvertTo-SecureString $c.ApiKey)
            )
        )
        return @{
            ApiKey = if ($decApi -eq "none") { "" } else { $decApi }
        }
    } catch { return $null }
}

# ============================================================
# BOZZE LOCALI
# ============================================================
function Load-Drafts {
    $list = New-Object System.Collections.ArrayList
    if (Test-Path $draftsFile) {
        try {
            $content = [System.IO.File]::ReadAllText($draftsFile, [System.Text.Encoding]::UTF8).Trim()
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                if (-not $content.StartsWith('[')) { $content = "[$content]" }
                $parsed = ConvertFrom-Json -InputObject $content
                if ($null -ne $parsed) {
                    foreach ($item in $parsed) {
                        if ($null -ne $item) {
                            $list.Add($item) | Out-Null
                        }
                    }
                }
            }
        } catch { }
    }
    # Forza restituzione ArrayList, non pipeline
    Write-Output -NoEnumerate $list
}

function Save-Draft {
    param($to, $cc, $subject, $body)
    # Carica lista esistente
    $drafts = Load-Drafts
    if ($null -eq $drafts) { $drafts = New-Object System.Collections.ArrayList }
    # Crea nuova bozza
    $draft = New-Object PSObject -Property @{
        Id      = [guid]::NewGuid().ToString()
        Date    = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        To      = "$to"
        Cc      = "$cc"
        Subject = "$subject"
        Body    = "$body"
    }
    $drafts.Add($draft) | Out-Null
    # Converti in array .NET puro e serializza
    $plain = New-Object System.Collections.ArrayList
    foreach ($d in $drafts) { $plain.Add($d) | Out-Null }
    $json = ConvertTo-Json -InputObject $plain -Depth 5
    # Se ancora non e' array JSON, forza wrapping
    if (-not $json.TrimStart().StartsWith('[')) { $json = "[$json]" }
    [System.IO.File]::WriteAllText($draftsFile, $json, [System.Text.Encoding]::UTF8)
}

function Delete-Draft {
    param($id)
    $drafts   = Load-Drafts
    $filtered = New-Object System.Collections.ArrayList
    if ($null -ne $drafts) {
        foreach ($d in $drafts) {
            if ($null -ne $d -and $d.Id -ne $id) { $filtered.Add($d) | Out-Null }
        }
    }
    $json = if ($filtered.Count -gt 0) { ConvertTo-Json -InputObject $filtered -Depth 5 } else { "[]" }
    if (-not $json.TrimStart().StartsWith('[')) { $json = "[$json]" }
    [System.IO.File]::WriteAllText($draftsFile, $json, [System.Text.Encoding]::UTF8)
}

# Aggiorna la listbox Bozze con i dati correnti
function Refresh-DraftsList {
    $script:lstDraftsRef.Items.Clear()
    $drafts = Load-Drafts
    $count  = 0
    if ($null -ne $drafts) {
        foreach ($d in $drafts) {
            if ($null -ne $d) {
                $script:lstDraftsRef.Items.Add("$($d.Date)  |  $($d.Subject)  ->  $($d.To)") | Out-Null
                $count++
            }
        }
    }
    if ($count -eq 0) {
        $script:lstDraftsRef.Items.Add("(nessuna bozza salvata)") | Out-Null
    }
}

# ============================================================
# OUTLOOK COM - INIZIALIZZAZIONE
# ============================================================
$script:outlookApp = $null

function Get-OutlookInstance {
    if ($script:outlookApp) { return $script:outlookApp }
    try {
        $script:outlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        return $script:outlookApp
    } catch {
        try {
            $script:outlookApp = New-Object -ComObject Outlook.Application
            return $script:outlookApp
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Impossibile avviare Outlook.`n`nAssicurati che:`n- Outlook sia installato`n- Outlook sia stato aperto almeno una volta`n- Il profilo sia configurato correttamente",
                "Errore Outlook",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $null
        }
    }
}

# ============================================================
# SELEZIONE ACCOUNT (multi-account: O365, Gmail, IMAP, ecc.)
# Nessuna credenziale richiesta o memorizzata: Outlook e' gia'
# autenticato su tutti gli account che ha configurato l'utente.
# ============================================================
function Get-OutlookAccounts {
    $accounts = [System.Collections.ArrayList]@()
    try {
        $ol = Get-OutlookInstance
        $ns = $ol.GetNamespace("MAPI")
        foreach ($acc in $ns.Session.Accounts) {
            try {
                $accounts.Add([PSCustomObject]@{
                    DisplayName = $acc.DisplayName
                    SmtpAddress = $acc.SmtpAddress
                }) | Out-Null
            } catch { continue }
        }
    } catch {}
    return $accounts
}

function Get-SelectedAccountObj {
    # Restituisce l'oggetto Account COM selezionato nelle impostazioni, o $null se predefinito
    try {
        $selected = if ($script:appSettings.ContainsKey('SelectedAccount')) { $script:appSettings.SelectedAccount } else { "" }
        if (-not $selected -or $selected -eq "" -or $selected -eq "(Account predefinito)") { return $null }
        $ol = Get-OutlookInstance
        $ns = $ol.GetNamespace("MAPI")
        foreach ($acc in $ns.Session.Accounts) {
            if ($acc.SmtpAddress -eq $selected -or $acc.DisplayName -eq $selected) { return $acc }
        }
    } catch {}
    return $null
}

function Get-SelectedAccountInbox {
    # Inbox dell'account scelto, o quella predefinita se non specificato (comportamento originale)
    try {
        $ol = Get-OutlookInstance
        $ns = $ol.GetNamespace("MAPI")
        $acc = Get-SelectedAccountObj
        if ($acc) {
            try { return $acc.DeliveryStore.GetDefaultFolder(6) } catch {}
        }
        return $ns.GetDefaultFolder(6)
    } catch { return $null }
}

# ============================================================
# FIX: RISOLUZIONE INDIRIZZI EMAIL PULITI (no Exchange DN)
# ============================================================
function Get-CleanEmail {
    param($mailItem)
    try {
        $email = $mailItem.SenderEmailAddress
        if ($email -match "^/O=" -or $email -match "^CN=") {
            try {
                $exchUser = $mailItem.Sender.GetExchangeUser()
                if ($exchUser) {
                    return $exchUser.PrimarySmtpAddress
                }
            } catch {}
            return $mailItem.SenderName
        }
        return $email
    } catch {
        return $mailItem.SenderName
    }
}

function Get-CleanRecipients {
    param($mailItem, $recipientType = 1)
    $addresses = [System.Collections.ArrayList]@()
    try {
        $recipients = $mailItem.Recipients
        for ($r = 1; $r -le $recipients.Count; $r++) {
            try {
                $recip = $recipients.Item($r)
                if ($recip.Type -eq $recipientType) {
                    $addr = $recip.Address
                    if ($addr -match "^/O=" -or $addr -match "^CN=") {
                        try {
                            $exchUser = $recip.AddressEntry.GetExchangeUser()
                            if ($exchUser) {
                                $addr = $exchUser.PrimarySmtpAddress
                            } else {
                                $addr = $recip.Name
                            }
                        } catch {
                            $addr = $recip.Name
                        }
                    }
                    if ($addr -and $addr -ne "") {
                        $addresses.Add($addr) | Out-Null
                    }
                }
            } catch { continue }
        }
    } catch {}
    return ($addresses -join "; ")
}

# ============================================================
# LETTURA EMAIL DA OUTLOOK COM (CON HTML E ALLEGATI)
# ============================================================
function Get-OutlookEmails {
    param($maxMessages = 30, $onProgress = $null)
    $emails = [System.Collections.ArrayList]@()
    try {
        $outlook   = Get-OutlookInstance
        if (-not $outlook) { return $emails }
        
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox     = Get-SelectedAccountInbox
        if (-not $inbox) { $inbox = $namespace.GetDefaultFolder(6) }
        $items     = $inbox.Items
        $items.Sort("[ReceivedTime]", $true)

        $count = [Math]::Min($maxMessages, $items.Count)
        for ($i = 1; $i -le $count; $i++) {
            try {
                $mail = $items.Item($i)
                
                $senderEmail = Get-CleanEmail $mail
                $toClean     = Get-CleanRecipients $mail 1
                $ccClean     = Get-CleanRecipients $mail 2
                
                # Estrai allegati visibili (Type != 5) e immagini inline CID (Type == 5)
                $attachList  = [System.Collections.ArrayList]@()
                $cidMap      = @{}   # CID -> percorso file locale

                if ($mail.Attachments.Count -gt 0) {
                    $imgDir = "$rootPath\img_cache\$($mail.EntryID -replace '[^a-zA-Z0-9]','_')"
                    for ($a = 1; $a -le $mail.Attachments.Count; $a++) {
                        try {
                            $att = $mail.Attachments.Item($a)
                            if ($att.Type -eq 5) {
                                # Immagine inline CID - estrai su disco
                                try {
                                    if (-not (Test-Path $imgDir)) {
                                        New-Item -ItemType Directory -Force -Path $imgDir | Out-Null
                                    }
                                    $fname   = $att.FileName
                                    $fpath   = Join-Path $imgDir $fname
                                    $att.SaveAsFile($fpath)
                                    # Il CID e' nel PropertyAccessor (PR_ATTACH_CONTENT_ID)
                                    try {
                                        $cid = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                        if ($cid) { $cidMap[$cid] = $fpath }
                                    } catch {}
                                    # Fallback: usa il nome file come CID
                                    if ($fname) { $cidMap[$fname] = $fpath }
                                } catch {}
                            } elseif ($att.FileName -ne "") {
                                $attachList.Add([PSCustomObject]@{
                                    Index = $a
                                    Name  = $att.FileName
                                    Size  = $att.Size
                                }) | Out-Null
                            }
                        } catch { continue }
                    }
                }

                # Riscrivi HTML: sostituisci cid:xxx con percorso file locale
                $htmlBody = $mail.HTMLBody
                if ($htmlBody -and $cidMap.Count -gt 0) {
                    foreach ($cid in $cidMap.Keys) {
                        $localPath = $cidMap[$cid].Replace('\','/')
                        $htmlBody  = $htmlBody -replace [regex]::Escape("cid:$cid"), "file:///$localPath"
                    }
                }

                $emails.Add([PSCustomObject]@{
                    Id           = $mail.EntryID
                    From         = $senderEmail
                    FromName     = $mail.SenderName
                    To           = $toClean
                    CC           = $ccClean
                    Subject      = if ($mail.Subject) { $mail.Subject } else { "(nessun oggetto)" }
                    Date         = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    Body         = $mail.Body
                    HTMLBody     = $htmlBody
                    Attachments  = $attachList
                    Unread       = $mail.UnRead
                    MailObject   = $mail
                }) | Out-Null
            } catch { continue }
            if ($onProgress) { try { & $onProgress $i $count } catch {} }
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match "MAPI_E|profile|Nessun profilo|no default profile|GetDefaultFolder|0x8004010F|0x80040111") {
            [System.Windows.Forms.MessageBox]::Show(
                "Outlook non ha nessun account email configurato.`n`n" +
                "Apri Outlook manualmente e configura almeno un account (Gmail, Outlook.com, Exchange, IMAP...) " +
                "prima di usare questo client. Non serve inserire credenziali qui: basta che l'account sia gia' attivo in Outlook.",
                "Nessun account configurato",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Errore lettura email:`n$errMsg", "Errore",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    return $emails
}

# ============================================================
# INVIO EMAIL TRAMITE OUTLOOK COM
# ============================================================
function Send-OutlookEmail {
    param($to, $cc, $subject, $body, $attachments)

    # Validazione indirizzi PRIMA di tentare l'invio - messaggio di errore chiaro
    # invece del generico errore COM di Outlook, e nessun rischio di perdere il testo
    function Test-EmailAddressFormat {
        param($addr)
        $addr = $addr.Trim()
        if ($addr -eq "") { return $true }   # vuoto va bene, verra' ignorato
        # Formato "Nome Cognome <email@dominio.it>" - estrai la parte email
        if ($addr -match '<(.+)>') { $addr = $Matches[1].Trim() }
        # Se non contiene @ e non e' vuoto, e' un nome GAL non risolto in indirizzo
        return ($addr -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
    }

    $badAddrs = [System.Collections.ArrayList]@()
    foreach ($addr in ($to -split ";")) {
        $a = $addr.Trim()
        if ($a -ne "" -and -not (Test-EmailAddressFormat $a)) { $badAddrs.Add("To: $a") | Out-Null }
    }
    if ($cc) {
        foreach ($addr in ($cc -split ";")) {
            $a = $addr.Trim()
            if ($a -ne "" -and -not (Test-EmailAddressFormat $a)) { $badAddrs.Add("Cc: $a") | Out-Null }
        }
    }
    if ($badAddrs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Uno o piu' indirizzi non sembrano validi (manca la @ o il dominio):`n`n" +
            ($badAddrs -join "`n") +
            "`n`nCorreggi l'indirizzo prima di inviare. Il testo della email non e' stato perso.",
            "Indirizzo non valido",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }

    try {
        $outlook = Get-OutlookInstance
        if (-not $outlook) { return $false }

        $mail             = $outlook.CreateItem(0)
        $mail.Subject     = $subject
        $mail.BodyFormat  = 2        # olFormatHTML - evita che i tag appaiano come testo
        $mail.HTMLBody    = $body    # corpo HTML dal WebBrowser editor

        # Se un account specifico e' stato scelto in Impostazioni, invia da quello
        $selAcc = Get-SelectedAccountObj
        if ($selAcc) { try { $mail.SendUsingAccount = $selAcc } catch {} }

        $to -split ";" | ForEach-Object {
            $addr = $_.Trim()
            if ($addr -ne "") {
                $recipient = $mail.Recipients.Add($addr)
                $recipient.Type = 1
            }
        }
        
        if ($cc) {
            $cc -split ";" | ForEach-Object {
                $addr = $_.Trim()
                if ($addr -ne "") {
                    $recipient = $mail.Recipients.Add($addr)
                    $recipient.Type = 2
                }
            }
        }

        $resolved = $mail.Recipients.ResolveAll()
        if (-not $resolved) {
            # Trova quali destinatari non si sono risolti per un messaggio piu' preciso
            $unresolved = [System.Collections.ArrayList]@()
            foreach ($r in $mail.Recipients) {
                if (-not $r.Resolved) { $unresolved.Add($r.Name) | Out-Null }
            }
            $detail = if ($unresolved.Count -gt 0) { "`n`nNon risolti: " + ($unresolved -join ", ") } else { "" }
            [System.Windows.Forms.MessageBox]::Show(
                "Impossibile verificare uno o piu' destinatari.$detail`n`nCorreggi l'indirizzo prima di inviare. Il testo della email non e' stato perso.",
                "Destinatario non risolto",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $false
        }

        foreach ($path in $attachments) {
            if (Test-Path $path) {
                $mail.Attachments.Add($path) | Out-Null
            }
        }

        $mail.Send()
        [System.Windows.Forms.MessageBox]::Show("Email inviata con successo!", "Successo",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore invio email:`n$_`n`nIl testo della email non e' stato perso: correggi e riprova.", "Errore Invio",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# ============================================================
# SALVA IN BOZZE OUTLOOK
# ============================================================
function Save-OutlookDraft {
    param($to, $cc, $subject, $body, $attachments)
    try {
        $outlook = Get-OutlookInstance
        if (-not $outlook) { return $false }
        
        $mail             = $outlook.CreateItem(0)
        $mail.Subject     = $subject
        $mail.Body        = $body

        $to -split ";" | ForEach-Object {
            $addr = $_.Trim()
            if ($addr -ne "") {
                $recipient = $mail.Recipients.Add($addr)
                $recipient.Type = 1
            }
        }
        if ($cc) {
            $cc -split ";" | ForEach-Object {
                $addr = $_.Trim()
                if ($addr -ne "") {
                    $recipient = $mail.Recipients.Add($addr)
                    $recipient.Type = 2
                }
            }
        }

        foreach ($path in $attachments) {
            if (Test-Path $path) {
                $mail.Attachments.Add($path) | Out-Null
            }
        }

        $mail.Save()
        [System.Windows.Forms.MessageBox]::Show("Bozza salvata in Outlook!", "Successo",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore salvataggio bozza:`n$_", "Errore Bozza",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# ============================================================
# CHIAMATA AI (Anthropic Claude)
# ============================================================
function Invoke-AI {
    param(
        $prompt,
        $apiKey,
        $replyContext = "",
        $language     = "Italiano",
        $tone         = "Professionale",
        $mode         = "new"
    )

    if (-not $apiKey) { throw "API Key non configurata. Impostala nel tab Impostazioni." }

    $modelId = "claude-haiku-4-5-20251001"
    try {
        if ($script:cmbModelRef -and $script:cmbModelRef.SelectedIndex -ge 0) {
            $sel = $script:cmbModelRef.SelectedItem.ToString()
            if ($sel -match "^(claude-\S+)") { $modelId = $Matches[1].Trim() }
        }
    } catch { }

    $toneInstr = switch ($tone) {
        "Formale"       { "Use a very formal and detached tone." }
        "Professionale" { "Use a professional but friendly tone." }
        "Amichevole"    { "Use a warm and direct tone, less formal." }
        "Informale"     { "Use an informal, conversational tone, as between colleagues." }
        "Conciso"       { "Be very concise: maximum 3-4 sentences, no filler." }
        default         { "Use a professional but friendly tone." }
    }

    $sigName = if ($script:appSettings.ContainsKey('SignatureName')) { $script:appSettings.SignatureName.Trim() } else { "" }
    $sigInstr = if ($sigName -ne "") {
        "CRITICAL: If a closing signature is appropriate, sign with the name '$sigName' (or a natural shorter/casual form of it if the tone is informal, e.g. just the first name). " +
        "NEVER sign with any other name found in reference material (original email, quoted thread, forwarded content) - those belong to other people."
    } else {
        "CRITICAL: Never sign the email with a name copied from any reference material (original email, quoted thread, forwarded content). Those names belong to other people. Only use a name if the user explicitly provides one to sign with."
    }

    $systemMsg = "You are an assistant that writes emails. " +
                 "IMPORTANT: You MUST write the email body in $language. " +
                 "Do NOT use any other language. $toneInstr " +
                 "CRITICAL: Always follow the user's explicit instructions. If the user gives specific directives (e.g. 'be brief', 'mention X', 'use formal tone', 'do not apologize'), those instructions take absolute priority over any other consideration. " +
                 "$sigInstr " +
                 "Write ONLY the email body text, without subject line or headers. " +
                 "Do not use typographic quotes (use ' and `") and do not use special dashes (use -)."

    # Aggiungi lo stile personale appreso, se presente e se l'uso e' abilitato
    try {
        $useStyle = if ($script:appSettings.ContainsKey('StyleUsageEnabled')) { [bool]$script:appSettings.StyleUsageEnabled } else { $true }
        if ($useStyle) {
            $styleProfile = Load-StyleProfile
            if ($styleProfile.Summary -and $styleProfile.Summary.Trim() -ne "") {
                $systemMsg += " Additionally, the user has a personal writing style learned from past edits: $($styleProfile.Summary) " +
                              "Apply the TONE, STRUCTURE and REGISTER described above, but ALWAYS write greetings, sign-offs and every word in $language. " +
                              "Never reuse literal words from another language even if the style description happens to mention them (e.g. if it mentions an English greeting or closing, use the natural $language equivalent instead, not the English word itself). " +
                              "Language and tone instructions above always take priority over the learned style when they conflict."
            }
        }
    } catch {}

    $userMsg = switch ($mode) {
        "reply" {
            "Your PRIMARY task is to follow the user's instructions below. " +
            "The original email is provided only as context/reference for tone and content - do NOT simply summarize or repeat it. " +
            "CRITICAL: The original email below may itself be a reply chain containing previous messages from OTHER people (colleagues, other senders). " +
            "Any names, signatures, or sign-offs that appear anywhere in that context belong to those other people. Follow the signature instructions given separately - never take a name from this context. " +
            "User instructions (FOLLOW THESE FIRST): $prompt`n`n" +
            "--- ORIGINAL EMAIL (context only - may contain a thread with other people's names, do not reuse them) ---`n$replyContext`n--- END ---`n`n" +
            "Now write the reply in $language following the user instructions above."
        }
        "forward" {
            "Your PRIMARY task is to follow the user's instructions below. " +
            "The forwarded email is provided only as context - do NOT reply to it, do NOT summarize it unless instructed. " +
            "CRITICAL: The forwarded content may contain names, signatures, or sign-offs from OTHER people (the original sender or previous participants in a thread). " +
            "Never copy or reuse any of those names. Follow the signature instructions given separately. " +
            "User instructions (FOLLOW THESE FIRST): $prompt`n`n" +
            "--- FORWARDED EMAIL (context only - may contain other people's names, do not reuse them) ---`n$replyContext`n--- END ---`n`n" +
            "Now write the introductory text in $language following the user instructions above."
        }
        default {
            "Your PRIMARY task is to follow the user's instructions below exactly. " +
            "User instructions (FOLLOW THESE): $prompt`n`n" +
            "Write the email body in $language following these instructions precisely."
        }
    }

    $payloadObj = @{
        model      = $modelId
        max_tokens = 1024
        system     = $systemMsg
        messages   = @(@{ role = "user"; content = $userMsg })
    }

    $payload   = $payloadObj | ConvertTo-Json -Depth 10
    $payload   = $payload -replace '\\u003c','<' -replace '\\u003e','>' `
                          -replace '\\u0027',"'" -replace '\\u0026','&'
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    $bodyBytes = $utf8NoBom.GetBytes($payload)

    try {
        $req = [System.Net.HttpWebRequest]::Create("https://api.anthropic.com/v1/messages")
        $req.Method = "POST"; $req.ContentType = "application/json; charset=utf-8"
        $req.ContentLength = $bodyBytes.Length
        $req.Headers.Add("x-api-key",         $apiKey)
        $req.Headers.Add("anthropic-version", "2023-06-01")
        $s = $req.GetRequestStream(); $s.Write($bodyBytes,0,$bodyBytes.Length); $s.Close()
        $resp   = $req.GetResponse()
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream(), $utf8NoBom)
        $json   = $reader.ReadToEnd(); $reader.Close(); $resp.Close()
        $result = $json | ConvertFrom-Json
        $text   = $result.content[0].text
        $text   = $text -replace [char]0x201C,'"' -replace [char]0x201D,'"' `
                        -replace [char]0x2018,"'"  -replace [char]0x2019,"'" `
                        -replace [char]0x2013,'-'  -replace [char]0x2014,'--' `
                        -replace [char]0x2026,'...'
        return $text

    } catch [System.Net.WebException] {
        $errBody = ""
        try {
            $es = $_.Exception.Response.GetResponseStream()
            $er = New-Object System.IO.StreamReader($es, $utf8NoBom)
            $errBody = $er.ReadToEnd(); $er.Close()
        } catch { }
        $errMsg = $errBody
        try {
            $o = $errBody | ConvertFrom-Json
            if ($o.error.message) { $errMsg = "[$($o.error.type)] $($o.error.message)" }
        } catch { }
        throw "Errore API Claude: $errMsg"
    } catch {
        throw "Errore connessione AI: $($_.Exception.Message)"
    }
}


# ============================================================
# VARIABILI GLOBALI
# ============================================================
$script:attachPaths    = [System.Collections.ArrayList]@()
$script:inboxEmails    = [System.Collections.ArrayList]@()
$script:selectedEmail  = $null
$script:pendingHTML    = ""
$script:config         = Load-Credentials
if (-not $script:config) {
    $script:config = @{ ApiKey = "" }
}

# ============================================================
# FORM PRINCIPALE
# ============================================================
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Copezzot Mail AI Assistant"
$form.Size          = New-Object System.Drawing.Size(1000, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor     = Parse-Color $script:appSettings.FormBgColor
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# Icona da file .ico accanto allo script
$iconPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "favicon.ico"
if (Test-Path $iconPath) { $form.Icon = New-Object System.Drawing.Icon($iconPath) }

# Logo Copezzot in base64 (incorporato nello script)
# Percorso logo - nella stessa cartella dello script
$script:logoPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "copezzot_logo.jpeg"
if (-not (Test-Path $script:logoPath)) {
    $script:logoPath = Join-Path $rootPath "copezzot_logo.jpeg"
}

# Carica immagine logo con alta qualita' (InterpolationMode = HighQualityBicubic)
function Load-Logo {
    param([int]$w=56, [int]$h=56)
    if (-not (Test-Path $script:logoPath)) { return $null }
    try {
        $src = [System.Drawing.Image]::FromFile($script:logoPath)
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.DrawImage($src, 0, 0, $w, $h)
        $g.Dispose(); $src.Dispose()
        return $bmp
    } catch { return $null }
}

$tab      = New-Object System.Windows.Forms.TabControl
$tab.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tab.BackColor = Parse-Color $script:appSettings.FormBgColor
$form.Controls.Add($tab)

# Logo Copezzot: pannello header giallo con logo a destra
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height    = 62
$pnlHeader.BackColor = Parse-Color $script:appSettings.HeaderColor
$form.Controls.Add($pnlHeader)

$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Size      = New-Object System.Drawing.Size(56, 56)
$picLogo.Location  = New-Object System.Drawing.Point(3, 3)
$picLogo.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$picLogo.BackColor = [System.Drawing.Color]::White
$picLogo.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$picLogo.Image = Load-Logo -w 56 -h 56
$pnlHeader.Controls.Add($picLogo)

# Riposiziona logo a destra con margine 3px ad ogni resize
$form.Add_Resize({
    $picLogo.Left = $pnlHeader.Width - $picLogo.Width - 3
})


# Titolo app nel header
$lblAppTitle = New-Object System.Windows.Forms.Label
$lblAppTitle.Text      = "Copezzot Mail AI Assistant"
$lblAppTitle.Location  = New-Object System.Drawing.Point(10, 18)
$lblAppTitle.Size      = New-Object System.Drawing.Size(500, 28)
$lblAppTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblAppTitle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$lblAppTitle.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblAppTitle)

# Bordo nero al form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.BackColor       = Parse-Color $script:appSettings.FormBgColor

# ============================================================
# HELPER UI
# ============================================================
function New-Label {
    param($text, $x, $y, $w = 70, $h = 22, $bold = $true)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, $h)
    $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, $(if($bold){[System.Drawing.FontStyle]::Bold}else{[System.Drawing.FontStyle]::Regular}))
    return $l
}

function New-TextBox {
    param($x, $y, $w, $h = 22, $multi = $false, $pwd = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y)
    $t.Size = New-Object System.Drawing.Size($w, $h)
    $t.Multiline = $multi
    if ($multi) { $t.ScrollBars = "Vertical" }
    if ($pwd) { $t.PasswordChar = "*" }
    return $t
}

function New-Btn {
    param($text, $x, $y, $w = 110, $h = 30, $color = "SteelBlue")
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.BackColor = [System.Drawing.Color]::$color
    $b.ForeColor = [System.Drawing.Color]::White
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    return $b
}

# ============================================================
# GAL AUTOCOMPLETE - PS5 compatible
# Mostra suggerimento in label dedicata (non tooltip)
# Cerca per nome, alias E indirizzo email nella GAL
# ============================================================
$script:galInstances = @{}
$script:galNextId    = 0

function Add-GalAutoComplete {
    param(
        [System.Windows.Forms.TextBox]$tb,
        [System.Windows.Forms.Control]$parent
    )

    $tb.Name = "gal_$($script:galNextId)"
    $script:galNextId++

    $popup = New-Object System.Windows.Forms.ListBox
    $popup.Visible     = $false
    $popup.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $popup.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
    $popup.BackColor   = [System.Drawing.Color]::White
    $popup.ForeColor   = [System.Drawing.Color]::FromArgb(0,70,150)
    $popup.ItemHeight  = 20
    $popup.Name        = "popup_$($tb.Name)"
    if ($parent) { $parent.Controls.Add($popup) }

    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = 280
    $timer.Tag      = $tb.Name

    $script:galInstances[$tb.Name] = @{
        Timer  = $timer
        Popup  = $popup
        Tb     = $tb
        Parent = $parent
    }

    $timer.Add_Tick({
        param($s, $e)
        $s.Stop()
        $inst = $script:galInstances[$s.Tag]
        if (-not $inst) { return }
        $ctrl = $inst.Tb
        $pp   = $inst.Popup

        $full = $ctrl.Text
        if (-not $full) { $pp.Visible = $false; return }

        $sep = $full.LastIndexOfAny(@([char]';', [char]','))
        $q   = if ($sep -ge 0) { $full.Substring($sep+1).Trim() } else { $full.Trim() }
        $minLen = if ($q -match '\.') { 2 } else { 3 }
        if ($q.Length -lt $minLen) { $pp.Visible = $false; return }

        # Cerca nella cache locale (istantaneo)
        $results = @(Search-Contacts -query $q -maxResults 8)

        # Fallback GAL solo se cache vuota
        if ($results.Count -eq 0) {
            try {
                $ol    = Get-OutlookInstance
                $ns    = $ol.GetNamespace("MAPI")
                $recip = $ns.CreateRecipient($q)
                if ($recip.Resolve()) {
                    $ae    = $recip.AddressEntry
                    $email = ""
                    try { $email = $ae.GetExchangeUser().PrimarySmtpAddress } catch {}
                    if (-not $email -and $ae.Address -match "@") { $email = $ae.Address }
                    $name = if ($ae.Name -and $ae.Name -ne $email) { $ae.Name } else { $email }
                    if ($email) {
                        Add-ContactToCache -email $email -name $name
                        $script:contactsCache = Load-Contacts
                        $results = @([PSCustomObject]@{ Email=$email; Name=$name; Display="$name <$email>" })
                    }
                }
            } catch {}
        }

        if ($results.Count -eq 0) { $pp.Visible = $false; return }

        $pp.Items.Clear()
        foreach ($r in $results) { $pp.Items.Add($r.Display) | Out-Null }

        # Posiziona popup sotto il campo
        if ($inst.Parent) {
            $screenPt = $ctrl.PointToScreen([System.Drawing.Point]::Empty)
            $parentPt = $inst.Parent.PointToClient($screenPt)
            $pp.Location = New-Object System.Drawing.Point($parentPt.X, ($parentPt.Y + $ctrl.Height + 2))
        }
        $pp.Width   = $ctrl.Width
        $pp.Height  = [Math]::Min($results.Count, 8) * 22 + 4
        $pp.Visible = $true
        $pp.BringToFront()
    })

    $tb.Add_TextChanged({
        param($s, $e)
        $inst = $script:galInstances[$s.Name]
        if (-not $inst) { return }
        $inst.Popup.Visible = $false
        $inst.Timer.Stop()
        $inst.Timer.Start()
    })

    $tb.Add_PreviewKeyDown({
        param($s, $e)
        $inst = $script:galInstances[$s.Name]
        if (-not $inst -or -not $inst.Popup.Visible) { return }
        if ($e.KeyCode -in @([System.Windows.Forms.Keys]::Tab,
                              [System.Windows.Forms.Keys]::Up,
                              [System.Windows.Forms.Keys]::Down,
                              [System.Windows.Forms.Keys]::Return)) {
            $e.IsInputKey = $true
        }
    })

    $tb.Add_KeyDown({
        param($s, $e)
        $inst = $script:galInstances[$s.Name]
        if (-not $inst) { return }
        $pp = $inst.Popup

        if ($pp.Visible) {
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Down) {
                $e.SuppressKeyPress = $true
                if ($pp.SelectedIndex -lt $pp.Items.Count-1) { $pp.SelectedIndex++ } else { $pp.SelectedIndex = 0 }
            } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Up) {
                $e.SuppressKeyPress = $true
                if ($pp.SelectedIndex -gt 0) { $pp.SelectedIndex-- } else { $pp.SelectedIndex = $pp.Items.Count-1 }
            } elseif ($e.KeyCode -in @([System.Windows.Forms.Keys]::Return, [System.Windows.Forms.Keys]::Tab)) {
                $e.SuppressKeyPress = $true; $e.Handled = $true
                $sel = if ($pp.SelectedIndex -ge 0) { $pp.SelectedItem } else { $pp.Items[0] }
                if ($sel) {
                    $full = $s.Text
                    $sep  = $full.LastIndexOfAny(@([char]';', [char]','))
                    $pre  = if ($sep -ge 0) { $full.Substring(0,$sep+1).TrimEnd() + " " } else { "" }
                    $s.Text = "$pre$sel"
                    $s.SelectionStart = $s.Text.Length
                }
                $pp.Visible = $false; $inst.Timer.Stop()
            } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $pp.Visible = $false; $inst.Timer.Stop()
            }
        }
    })

    $popup.Add_Click({
        param($s, $e)
        $key  = $s.Name -replace "^popup_",""
        $inst = $script:galInstances[$key]
        if (-not $inst -or $s.SelectedIndex -lt 0) { return }
        $ctrl = $inst.Tb
        $sel  = $s.SelectedItem.ToString()
        $full = $ctrl.Text
        $sep  = $full.LastIndexOfAny(@([char]';', [char]','))
        $pre  = if ($sep -ge 0) { $full.Substring(0,$sep+1).TrimEnd() + " " } else { "" }
        $ctrl.Text = "$pre$sel"
        $ctrl.SelectionStart = $ctrl.Text.Length
        $s.Visible = $false; $inst.Timer.Stop()
        $ctrl.Focus()
    })

    $tb.Add_Leave({
        param($s, $e)
        $inst = $script:galInstances[$s.Name]
        if (-not $inst) { return }
        $closeT = New-Object System.Windows.Forms.Timer
        $closeT.Interval = 200; $closeT.Tag = $s.Name
        $closeT.Add_Tick({
            param($cs,$ce); $cs.Stop()
            $ci = $script:galInstances[$cs.Tag]
            if ($ci) { $ci.Popup.Visible = $false }
        })
        $closeT.Start()
    })
}

# TAB 1 - NUOVA EMAIL
# ============================================================
$tabCompose = New-Object System.Windows.Forms.TabPage
$tabCompose.Text = "Nuova Email"
$tab.Controls.Add($tabCompose)

$lblTo = New-Label "To:" 10 15
$tabCompose.Controls.Add($lblTo)
$txtTo = New-TextBox 80 12 870
$tabCompose.Controls.Add($txtTo)
Add-GalAutoComplete $txtTo $tabCompose

$lblCc = New-Label "CC:" 10 45
$tabCompose.Controls.Add($lblCc)
$txtCc = New-TextBox 80 42 870
$tabCompose.Controls.Add($txtCc)
Add-GalAutoComplete $txtCc $tabCompose

$lblSubj = New-Label "Oggetto:" 10 75
$tabCompose.Controls.Add($lblSubj)
$txtSubj = New-TextBox 80 72 870
$tabCompose.Controls.Add($txtSubj)

$lblAttach = New-Label "Allegati:" 10 105
$tabCompose.Controls.Add($lblAttach)
$lstAttach = New-Object System.Windows.Forms.ListBox
$lstAttach.Location = New-Object System.Drawing.Point(80, 102)
$lstAttach.Size = New-Object System.Drawing.Size(640, 43)
$tabCompose.Controls.Add($lstAttach)

$btnAddAtt = New-Btn "Aggiungi" 730 102 100 22 "SeaGreen"
$tabCompose.Controls.Add($btnAddAtt)
$btnRemAtt = New-Btn "Rimuovi" 730 128 100 22 "Tomato"
$tabCompose.Controls.Add($btnRemAtt)

$lblBody = New-Label "Corpo:" 10 155 60 22
$tabCompose.Controls.Add($lblBody)

# -- TOOLBAR RICH TEXT (2 righe) -----------------------------
$pnlToolbar = New-Object System.Windows.Forms.Panel
$pnlToolbar.Location  = New-Object System.Drawing.Point(10, 175)
$pnlToolbar.Size      = New-Object System.Drawing.Size(930, 62)
$pnlToolbar.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 240)
$tabCompose.Controls.Add($pnlToolbar)

function New-TBtn { param($label,$x,$y=3,$w=28,$h=26,$fnt=$null)
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$label; $b.Location=New-Object System.Drawing.Point($x,$y)
    $b.Size=New-Object System.Drawing.Size($w,$h); $b.FlatStyle="Flat"
    $b.BackColor=[System.Drawing.Color]::White
    if($fnt){$b.Font=$fnt} else {$b.Font=New-Object System.Drawing.Font("Segoe UI",8)}
    return $b
}

# RIGA 1 - formattazione carattere
$tbBold   = New-TBtn "B"  4  3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold))
$tbItal   = New-TBtn "I"  35 3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Italic))
$tbUnder  = New-TBtn "U"  66 3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Underline))
$tbStrike = New-TBtn "S"  97 3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Strikeout))

$sep1=New-Object System.Windows.Forms.Label; $sep1.Location=New-Object System.Drawing.Point(130,4)
$sep1.Size=New-Object System.Drawing.Size(1,20); $sep1.BorderStyle="FixedSingle"; $pnlToolbar.Controls.Add($sep1)

$cmbFont = New-Object System.Windows.Forms.ComboBox
$cmbFont.Location=New-Object System.Drawing.Point(135,4); $cmbFont.Size=New-Object System.Drawing.Size(135,24)
$cmbFont.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbFont.Font=New-Object System.Drawing.Font("Segoe UI",8)
@("Aptos","Segoe UI","Arial","Calibri","Times New Roman","Courier New","Verdana","Georgia","Tahoma") |
    ForEach-Object { $cmbFont.Items.Add($_) | Out-Null }
$cmbFont.SelectedIndex=0   # Aptos

$cmbSize = New-Object System.Windows.Forms.ComboBox
$cmbSize.Location=New-Object System.Drawing.Point(274,4); $cmbSize.Size=New-Object System.Drawing.Size(54,24)
$cmbSize.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbSize.Font=New-Object System.Drawing.Font("Segoe UI",8)
@("8","9","10","11","12","14","16","18","20","24","28","32","36") |
    ForEach-Object { $cmbSize.Items.Add($_) | Out-Null }
$cmbSize.SelectedIndex = 4   # 12pt

$sep2=New-Object System.Windows.Forms.Label; $sep2.Location=New-Object System.Drawing.Point(333,4)
$sep2.Size=New-Object System.Drawing.Size(1,20); $sep2.BorderStyle="FixedSingle"; $pnlToolbar.Controls.Add($sep2)

$btnFgColor = New-TBtn "A" 337 3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold))
$btnFgColor.ForeColor=[System.Drawing.Color]::Red
$btnBgColor = New-TBtn "ab" 368 3 32 26; $btnBgColor.BackColor=[System.Drawing.Color]::Yellow

$sep3=New-Object System.Windows.Forms.Label; $sep3.Location=New-Object System.Drawing.Point(405,4)
$sep3.Size=New-Object System.Drawing.Size(1,20); $sep3.BorderStyle="FixedSingle"; $pnlToolbar.Controls.Add($sep3)

$lblLS=New-Object System.Windows.Forms.Label; $lblLS.Text="Interl:"
$lblLS.Location=New-Object System.Drawing.Point(410,7); $lblLS.Size=New-Object System.Drawing.Size(38,16)
$lblLS.Font=New-Object System.Drawing.Font("Segoe UI",8); $pnlToolbar.Controls.Add($lblLS)

$cmbLS = New-Object System.Windows.Forms.ComboBox
$cmbLS.Location=New-Object System.Drawing.Point(450,4); $cmbLS.Size=New-Object System.Drawing.Size(55,24)
$cmbLS.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbLS.Font=New-Object System.Drawing.Font("Segoe UI",8)
@("0","0.5","0.8","1.0","1.15","1.5","2.0","2.5","3.0") | ForEach-Object { $cmbLS.Items.Add($_) | Out-Null }
$cmbLS.SelectedIndex=4   # 1.15 default

# GENERA AI - destra riga 1
$btnAI = New-Object System.Windows.Forms.Button
$btnAI.Text="Genera AI"; $btnAI.Location=New-Object System.Drawing.Point(720,3)
$btnAI.Size=New-Object System.Drawing.Size(205,26)
$btnAI.BackColor=[System.Drawing.Color]::MediumPurple
$btnAI.ForeColor=[System.Drawing.Color]::White
$btnAI.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnAI.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

# RIGA 2 - allineamento + elenchi + rientro
$tbAlL   = New-TBtn "L"  4   33 28 26; $tbAlL.ForeColor  = [System.Drawing.Color]::DarkBlue
$tbAlC   = New-TBtn "C"  35  33 28 26; $tbAlC.ForeColor  = [System.Drawing.Color]::DarkBlue
$tbAlR   = New-TBtn "R"  66  33 28 26; $tbAlR.ForeColor  = [System.Drawing.Color]::DarkBlue
$tbAlJ   = New-TBtn "J"  97  33 28 26; $tbAlJ.ForeColor  = [System.Drawing.Color]::DarkBlue

$sep4=New-Object System.Windows.Forms.Label; $sep4.Location=New-Object System.Drawing.Point(130,34)
$sep4.Size=New-Object System.Drawing.Size(1,20); $sep4.BorderStyle="FixedSingle"; $pnlToolbar.Controls.Add($sep4)

$tbUl    = New-TBtn "ul" 135  33 30 26
$tbOl    = New-TBtn "ol" 168  33 30 26

$sep5=New-Object System.Windows.Forms.Label; $sep5.Location=New-Object System.Drawing.Point(203,34)
$sep5.Size=New-Object System.Drawing.Size(1,20); $sep5.BorderStyle="FixedSingle"; $pnlToolbar.Controls.Add($sep5)

$tbIndIn  = New-TBtn ">>" 208  33 32 26
$tbIndOut = New-TBtn "<<" 243  33 32 26

# Tooltip
$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($tbBold,"Grassetto"); $tip.SetToolTip($tbItal,"Corsivo")
$tip.SetToolTip($tbUnder,"Sottolineato"); $tip.SetToolTip($tbStrike,"Barrato")
$tip.SetToolTip($cmbFont,"Tipo carattere"); $tip.SetToolTip($cmbSize,"Dimensione")
$tip.SetToolTip($btnFgColor,"Colore testo"); $tip.SetToolTip($btnBgColor,"Colore sfondo/evidenziazione")
$tip.SetToolTip($cmbLS,"Interlinea")
$tip.SetToolTip($tbAlL,"Allinea sinistra"); $tip.SetToolTip($tbAlC,"Centra")
$tip.SetToolTip($tbAlR,"Allinea destra"); $tip.SetToolTip($tbAlJ,"Giustifica")
$tip.SetToolTip($tbUl,"Elenco puntato"); $tip.SetToolTip($tbOl,"Elenco numerato")
$tip.SetToolTip($tbIndIn,"Aumenta rientro"); $tip.SetToolTip($tbIndOut,"Diminuisci rientro")

@($tbBold,$tbItal,$tbUnder,$tbStrike,$cmbFont,$cmbSize,$btnFgColor,$btnBgColor,$cmbLS,$btnAI,
  $tbAlL,$tbAlC,$tbAlR,$tbAlJ,$tbUl,$tbOl,$tbIndIn,$tbIndOut) |
    ForEach-Object { $pnlToolbar.Controls.Add($_) }

# -- WebBrowser editor ----------------------------------------
# Bottoni azione nella toolbar riga 2 (destra) - sempre visibili
$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text      = "INVIA"
$btnSend.Location  = New-Object System.Drawing.Point(530, 33)
$btnSend.Size      = New-Object System.Drawing.Size(110, 26)
$btnSend.BackColor = [System.Drawing.Color]::SeaGreen
$btnSend.ForeColor = [System.Drawing.Color]::White
$btnSend.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSend.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlToolbar.Controls.Add($btnSend)

$btnSaveDraft = New-Object System.Windows.Forms.Button
$btnSaveDraft.Text      = "Salva Bozza"
$btnSaveDraft.Location  = New-Object System.Drawing.Point(645, 33)
$btnSaveDraft.Size      = New-Object System.Drawing.Size(110, 26)
$btnSaveDraft.BackColor = [System.Drawing.Color]::SteelBlue
$btnSaveDraft.ForeColor = [System.Drawing.Color]::White
$btnSaveDraft.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveDraft.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlToolbar.Controls.Add($btnSaveDraft)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text      = "Pulisci"
$btnClear.Location  = New-Object System.Drawing.Point(760, 33)
$btnClear.Size      = New-Object System.Drawing.Size(80, 26)
$btnClear.BackColor = [System.Drawing.Color]::Tomato
$btnClear.ForeColor = [System.Drawing.Color]::White
$btnClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClear.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlToolbar.Controls.Add($btnClear)

# WebBrowser occupa tutto lo spazio rimasto sotto la toolbar
$webCompose = New-Object System.Windows.Forms.WebBrowser
$webCompose.Location = New-Object System.Drawing.Point(10, 240)
$webCompose.Size     = New-Object System.Drawing.Size(930, 380)
$webCompose.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor
                       [System.Windows.Forms.AnchorStyles]::Left -bor
                       [System.Windows.Forms.AnchorStyles]::Right -bor
                       [System.Windows.Forms.AnchorStyles]::Bottom
$webCompose.IsWebBrowserContextMenuEnabled = $true
$webCompose.WebBrowserShortcutsEnabled     = $true
$webCompose.ScriptErrorsSuppressed         = $true
$tabCompose.Controls.Add($webCompose)
$script:webCompose = $webCompose


# HTML iniziale editor con stili di default
$initHtml = @"
<html>
<head>
<style>
  body {
    font-family: Aptos, 'Segoe UI', sans-serif;
    font-size: 12pt;
    line-height: 1.15;
    margin: 8px;
    color: #222;
    outline: none;
  }
</style>
</head>
<body contenteditable="true" id="mailbody"><br></body>
</html>
"@
$webCompose.DocumentText = $initHtml

$webCompose.Add_DocumentCompleted({
    try {
        $b = $script:webCompose.Document.Body
        if ($b) {
            $b.SetAttribute("contenteditable","true")
            $b.Id = "mailbody"
        }
    } catch { }
})

# Helper: esegui comando execCommand
function Exec-EditCmd { param($cmd,$val="")
    if($script:webCompose.Document){ $script:webCompose.Document.ExecCommand($cmd,$false,$val) | Out-Null }
}

# Helper: leggi HTML dal corpo
function Get-ComposeHtml {
    try {
        $body = $script:webCompose.Document.Body
        if($body){ return $body.InnerHtml } else { return "" }
    } catch { return "" }
}

# Helper: imposta HTML nel corpo preservando stile body
function Set-ComposeHtml { param($html)
    try {
        $body = $script:webCompose.Document.Body
        if ($body) {
            $body.InnerHtml = if ($html) { $html } else { "<br>" }
            $body.SetAttribute("contenteditable","true")
        }
    } catch {
        # Fallback: ricarica documento
        $lh = try { $cmbLS.SelectedItem.ToString() } catch { "1.15" }
        $escaped = if ($html) { $html } else { "<br>" }
        $h = "<html><head><style>body{font-family:Aptos,'Segoe UI',sans-serif;font-size:12pt;line-height:$lh;margin:8px;color:#222;}</style></head><body contenteditable='true'>$escaped</body></html>"
        $script:webCompose.DocumentText = $h
    }
}

# Helper: testo plain dal corpo
function Get-ComposePlain {
    try {
        $body = $script:webCompose.Document.Body
        if($body){ return $body.InnerText } else { return "" }
    } catch { return "" }
}

# Toolbar handlers - formattazione
$tbBold.Add_Click({   Exec-EditCmd "bold" })
$tbItal.Add_Click({   Exec-EditCmd "italic" })
$tbUnder.Add_Click({  Exec-EditCmd "underline" })
$tbStrike.Add_Click({ Exec-EditCmd "strikethrough" })

# Font family
$cmbFont.Add_SelectedIndexChanged({
    Exec-EditCmd "fontName" $cmbFont.SelectedItem.ToString()
})

# Font size (execCommand fontsize usa 1-7, mappiamo i pt)
$cmbSize.Add_SelectedIndexChanged({
    $pt = [int]$cmbSize.SelectedItem.ToString()
    $fs = switch ($pt) {
        {$_ -le 8}  { 1 } {$_ -le 10} { 2 } {$_ -le 12} { 3 }
        {$_ -le 14} { 4 } {$_ -le 18} { 5 } {$_ -le 24} { 6 }
        default     { 7 }
    }
    Exec-EditCmd "fontSize" $fs
})

# Colore testo
$btnFgColor.Add_Click({
    $dlg = New-Object System.Windows.Forms.ColorDialog
    $dlg.FullOpen = $true
    if ($dlg.ShowDialog() -eq "OK") {
        $hex = "#{0:X2}{1:X2}{2:X2}" -f $dlg.Color.R,$dlg.Color.G,$dlg.Color.B
        $script:fgColor = $hex
        $btnFgColor.ForeColor = $dlg.Color
        Exec-EditCmd "foreColor" $hex
    }
})

# Colore evidenziazione
$btnBgColor.Add_Click({
    $dlg = New-Object System.Windows.Forms.ColorDialog
    $dlg.FullOpen = $true
    if ($dlg.ShowDialog() -eq "OK") {
        $hex = "#{0:X2}{1:X2}{2:X2}" -f $dlg.Color.R,$dlg.Color.G,$dlg.Color.B
        $script:bgColor = $hex
        $btnBgColor.BackColor = $dlg.Color
        Exec-EditCmd "hiliteColor" $hex
    }
})

# Interlinea - applica line-height CSS al body del documento
$cmbLS.Add_SelectedIndexChanged({
    $lh = $cmbLS.SelectedItem.ToString()
    try {
        $script:webCompose.Document.Body.Style = "line-height:$lh; font-family:'Segoe UI',sans-serif; font-size:10pt; margin:8px; color:#222;"
    } catch { }
})

# Allineamento
$tbAlL.Add_Click({  Exec-EditCmd "justifyLeft" })
$tbAlC.Add_Click({  Exec-EditCmd "justifyCenter" })
$tbAlR.Add_Click({  Exec-EditCmd "justifyRight" })
$tbAlJ.Add_Click({  Exec-EditCmd "justifyFull" })

# Elenchi
$tbUl.Add_Click({     Exec-EditCmd "insertUnorderedList" })
$tbOl.Add_Click({     Exec-EditCmd "insertOrderedList" })

# Rientro
$tbIndIn.Add_Click({  Exec-EditCmd "indent" })
$tbIndOut.Add_Click({ Exec-EditCmd "outdent" })

# Bottoni azione
# (bottoni INVIA / Salva Bozza / Pulisci sono nella toolbar)

# ============================================================
# TAB 2 - POSTA IN ARRIVO
# ============================================================
$tabInbox = New-Object System.Windows.Forms.TabPage
$tabInbox.Text = "Posta in Arrivo"
$tab.Controls.Add($tabInbox)

$btnRefresh = New-Btn "Aggiorna" 10 10 120 30 "SteelBlue"
$tabInbox.Controls.Add($btnRefresh)
$lblStatus = New-Label "" 450 15 500 20 $false
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$tabInbox.Controls.Add($lblStatus)

$progInboxRefresh = New-Object System.Windows.Forms.ProgressBar
$progInboxRefresh.Location = New-Object System.Drawing.Point(140, 10)
$progInboxRefresh.Size     = New-Object System.Drawing.Size(300, 20)
$progInboxRefresh.Minimum  = 0
$progInboxRefresh.Maximum  = 100
$progInboxRefresh.Value    = 0
$progInboxRefresh.Visible  = $false
$tabInbox.Controls.Add($progInboxRefresh)
$script:progInboxRef = $progInboxRefresh

$lstEmails = New-Object System.Windows.Forms.ListView
$lstEmails.Location = New-Object System.Drawing.Point(10, 45)
$lstEmails.Size = New-Object System.Drawing.Size(940, 120)
$lstEmails.View = [System.Windows.Forms.View]::Details
$lstEmails.FullRowSelect = $true
$lstEmails.GridLines = $true
$lstEmails.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Colonne
$col1 = New-Object System.Windows.Forms.ColumnHeader
$col1.Text = "Mittente"; $col1.Width = 200
$col2 = New-Object System.Windows.Forms.ColumnHeader
$col2.Text = "Oggetto"; $col2.Width = 380
$col3 = New-Object System.Windows.Forms.ColumnHeader
$col3.Text = "Data"; $col3.Width = 130
$col4 = New-Object System.Windows.Forms.ColumnHeader
$col4.Text = "Stato"; $col4.Width = 100

$lstEmails.Columns.AddRange(@($col1,$col2,$col3,$col4))

# Ordinamento nativo al click header
$lstEmails.Add_ColumnClick({
    param($s,$e)
    $lstEmails.Sorting = if ($lstEmails.Sorting -eq 
        [System.Windows.Forms.SortOrder]::Ascending) {
        [System.Windows.Forms.SortOrder]::Descending
    } else {
        [System.Windows.Forms.SortOrder]::Ascending
    }
    $lstEmails.Sort()
})

$tabInbox.Controls.Add($lstEmails)

# Campi lettura email
$lblReadSubj = New-Label "Oggetto:" 10 175
$tabInbox.Controls.Add($lblReadSubj)
$txtReadSubj = New-TextBox 80 172 870 22
$txtReadSubj.ReadOnly = $true
$tabInbox.Controls.Add($txtReadSubj)

$lblReadFrom = New-Label "Da:" 10 205
$tabInbox.Controls.Add($lblReadFrom)
$txtReadFrom = New-TextBox 80 202 870 22
$txtReadFrom.ReadOnly = $true
$tabInbox.Controls.Add($txtReadFrom)

# Bottoni Reply (ORA VISIBILI E FUNZIONANTI)
$btnReply = New-Btn "Rispondi" 10 235 110 30 "MediumPurple"
$tabInbox.Controls.Add($btnReply)
$btnReplyAll = New-Btn "Rispondi a tutti" 130 235 130 30 "SteelBlue"
$tabInbox.Controls.Add($btnReplyAll)
$btnForward = New-Btn "Inoltra" 270 235 100 30 "DarkSeaGreen"
$tabInbox.Controls.Add($btnForward)

# Panel per WebBrowser (corpo email HTML)
$pnlBody = New-Object System.Windows.Forms.Panel
$pnlBody.Location = New-Object System.Drawing.Point(10, 275)
$btnLoadImages = New-Object System.Windows.Forms.Button
$btnLoadImages.Text      = "Scarica immagini"
$btnLoadImages.Location  = New-Object System.Drawing.Point(795, 250)
$btnLoadImages.Size      = New-Object System.Drawing.Size(155, 24)
$btnLoadImages.BackColor = [System.Drawing.Color]::FromArgb(240,240,150)
$btnLoadImages.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadImages.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$btnLoadImages.Visible   = $false
$tabInbox.Controls.Add($btnLoadImages)
$script:btnLoadImagesRef = $btnLoadImages

$pnlBody.Size = New-Object System.Drawing.Size(940, 420)
$pnlBody.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tabInbox.Controls.Add($pnlBody)

$webBody = New-Object System.Windows.Forms.WebBrowser
$webBody.Dock = [System.Windows.Forms.DockStyle]::Fill
$webBody.ScrollBarsEnabled = $true
$webBody.IsWebBrowserContextMenuEnabled = $false
$webBody.AllowWebBrowserDrop = $false
$webBody.ScriptErrorsSuppressed = $true
$webBody.WebBrowserShortcutsEnabled = $false
$pnlBody.Controls.Add($webBody)

# File HTML temporaneo per rendering con immagini inline (CID)
$script:tempHtmlPath = "$rootPath\preview_temp.html"
$script:currentEmailHtmlOrig = ""   # HTML originale (con immagini remote) dell'email selezionata

# Blocca immagini/risorse remote (http/https) di default - come "download immagini"
# di Outlook. Sostituisce src remoti con un placeholder, senza toccare le CID (cid:/file:///)
function Block-RemoteImages {
    param($html)
    if (-not $html) { return $html }
    # Sostituisce src="http(s)://..." con un src trasparente 1x1 - le richieste di rete
    # per immagini/CSS/iframe remoti sono la causa dei blocchi UI di diversi secondi
    $blocked = $html -replace '(?i)(<img[^>]+src\s*=\s*["''])https?://[^"'']+(["''])', '${1}data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBTAA7${2}'
    $blocked = $blocked -replace '(?i)<link[^>]+rel\s*=\s*["'']stylesheet["''][^>]*>', ''  # rimuovi CSS esterni
    return $blocked
}

# Timer per refresh WebBrowser
$script:webTimer = New-Object System.Windows.Forms.Timer
$script:webTimer.Interval = 100

$script:webTimer.Add_Tick({
    $script:webTimer.Stop()
    if ($script:pendingHTML -ne "") {
        try {
            # Scrivi su file temporaneo e naviga: permette immagini CID e risorse inline
            [System.IO.File]::WriteAllText($script:tempHtmlPath, $script:pendingHTML, [System.Text.Encoding]::UTF8)
            $webBody.Navigate("file:///$($script:tempHtmlPath.Replace('\','/'))")
            $script:pendingHTML = ""
        } catch {
            # Fallback a DocumentText se il file non funziona
            try { $webBody.DocumentText = $script:pendingHTML } catch {}
            $script:pendingHTML = ""
        }
    }
})

$btnLoadImages.Add_Click({
    if ($script:currentEmailHtmlOrig -ne "") {
        $script:pendingHTML = $script:currentEmailHtmlOrig
        $script:webTimer.Stop()
        $webBody.DocumentText = "<html><body></body></html>"
        $script:webTimer.Start()
        $btnLoadImages.Visible = $false
    }
})

# Allegati lista
$lblAttachRead = New-Label "Allegati:" 10 700
$tabInbox.Controls.Add($lblAttachRead)
$lstAttachRead = New-Object System.Windows.Forms.ListBox
$lstAttachRead.Location = New-Object System.Drawing.Point(80, 697)
$lstAttachRead.Size = New-Object System.Drawing.Size(600, 30)
$tabInbox.Controls.Add($lstAttachRead)

$btnSaveAttach = New-Btn "Salva Allegato" 690 695 120 30 "SteelBlue"
$tabInbox.Controls.Add($btnSaveAttach)

# ============================================================
# TAB 3 - BOZZE
# ============================================================
$tabDrafts = New-Object System.Windows.Forms.TabPage
$tabDrafts.Text = "Bozze"
$tabDrafts.AutoScroll = $true
$tab.Controls.Add($tabDrafts)

$lstDrafts = New-Object System.Windows.Forms.ListBox
$lstDrafts.Location = New-Object System.Drawing.Point(10, 10)
$lstDrafts.Size = New-Object System.Drawing.Size(940, 500)
$lstDrafts.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tabDrafts.Controls.Add($lstDrafts)
$script:lstDraftsRef = $lstDrafts   # riferimento script scope per Refresh-DraftsList

$btnLoadDraft = New-Btn "Carica in Compose" 10 520 140 35 "SteelBlue"
$tabDrafts.Controls.Add($btnLoadDraft)
$btnDeleteDraft = New-Btn "Elimina" 160 520 100 35 "Tomato"
$tabDrafts.Controls.Add($btnDeleteDraft)
$btnResetDrafts = New-Btn "Resetta lista" 270 520 120 35 "DarkGray"
$tabDrafts.Controls.Add($btnResetDrafts)

# ============================================================
# TAB 4 - IMPOSTAZIONI
# ============================================================
$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "Impostazioni"
$tabSettings.AutoScroll = $true
$tab.Controls.Add($tabSettings)

# ============================================================
# TAB - IL MIO STILE (apprendimento stile personale)
# ============================================================
$tabStyle = New-Object System.Windows.Forms.TabPage
$tabStyle.Text = "Il mio stile"
$tabStyle.AutoScroll = $true
$tab.Controls.Add($tabStyle)

$lblStyleTitle = New-Object System.Windows.Forms.Label
$lblStyleTitle.Text = "Come Claude sta imparando il tuo stile di scrittura"
$lblStyleTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblStyleTitle.Size     = New-Object System.Drawing.Size(700, 24)
$lblStyleTitle.Font     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$tabStyle.Controls.Add($lblStyleTitle)

$lblStyleDesc = New-Object System.Windows.Forms.Label
$lblStyleDesc.Text = "Ogni volta che invii una email dopo aver modificato una bozza generata dall'AI, Claude confronta le due versioni e aggiorna il profilo di stile qui sotto. Il profilo viene poi usato per le generazioni future."
$lblStyleDesc.Location = New-Object System.Drawing.Point(10, 36)
$lblStyleDesc.Size     = New-Object System.Drawing.Size(920, 40)
$lblStyleDesc.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblStyleDesc.ForeColor= [System.Drawing.Color]::Gray
$tabStyle.Controls.Add($lblStyleDesc)

# -- Interruttore abilita/disabilita apprendimento stile --
$grpToggle = New-Object System.Windows.Forms.GroupBox
$grpToggle.Text     = "Apprendimento e utilizzo stile"
$grpToggle.Location = New-Object System.Drawing.Point(10, 78)
$grpToggle.Size     = New-Object System.Drawing.Size(920, 105)
$tabStyle.Controls.Add($grpToggle)

$chkStyleEnabled = New-Object System.Windows.Forms.CheckBox
$chkStyleEnabled.Text     = "Abilita analisi automatica dello stile ad ogni invio (apprendimento)"
$chkStyleEnabled.Location = New-Object System.Drawing.Point(10, 22)
$chkStyleEnabled.Size     = New-Object System.Drawing.Size(500, 24)
$chkStyleEnabled.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$chkStyleEnabled.Checked  = [bool]$script:appSettings.StyleLearningEnabled
$grpToggle.Controls.Add($chkStyleEnabled)

$chkStyleUsage = New-Object System.Windows.Forms.CheckBox
$chkStyleUsage.Text     = "Usa il profilo di stile salvato nelle generazioni AI (utilizzo)"
$chkStyleUsage.Location = New-Object System.Drawing.Point(10, 48)
$chkStyleUsage.Size     = New-Object System.Drawing.Size(500, 24)
$chkStyleUsage.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$chkStyleUsage.Checked  = if ($script:appSettings.ContainsKey('StyleUsageEnabled')) { [bool]$script:appSettings.StyleUsageEnabled } else { $true }
$grpToggle.Controls.Add($chkStyleUsage)

$lblToggleNote = New-Object System.Windows.Forms.Label
$lblToggleNote.Text = "Apprendimento: analizza le tue modifiche dopo ogni invio (piccolo costo Claude Haiku, solo se il testo e' stato modificato). " +
                      "Utilizzo: applica il profilo gia' salvato alle nuove generazioni, anche se l'apprendimento e' disattivato. " +
                      "Il profilo risiede SOLO in locale (file style_<tua-email>.json) e non viene mai condiviso o riutilizzato da Anthropic al di fuori della singola chiamata di analisi."
$lblToggleNote.Location = New-Object System.Drawing.Point(10, 74)
$lblToggleNote.Size     = New-Object System.Drawing.Size(900, 28)
$lblToggleNote.AutoSize = $false
$lblToggleNote.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
$lblToggleNote.ForeColor= [System.Drawing.Color]::FromArgb(80,80,80)
$grpToggle.Controls.Add($lblToggleNote)

$chkStyleEnabled.Add_CheckedChanged({
    $script:appSettings.StyleLearningEnabled = $chkStyleEnabled.Checked
    Save-AppSettings $script:appSettings
})
$chkStyleUsage.Add_CheckedChanged({
    $script:appSettings.StyleUsageEnabled = $chkStyleUsage.Checked
    Save-AppSettings $script:appSettings
})

$lblStyleEmail = New-Object System.Windows.Forms.Label
$lblStyleEmail.Location = New-Object System.Drawing.Point(10, 195)
$lblStyleEmail.Size     = New-Object System.Drawing.Size(920, 20)
$lblStyleEmail.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
$lblStyleEmail.ForeColor= [System.Drawing.Color]::DarkSlateGray
$tabStyle.Controls.Add($lblStyleEmail)

$lblCurrentStyle = New-Object System.Windows.Forms.Label
$lblCurrentStyle.Text = "Profilo di stile attuale:"
$lblCurrentStyle.Location = New-Object System.Drawing.Point(10, 220)
$lblCurrentStyle.Size     = New-Object System.Drawing.Size(300, 20)
$lblCurrentStyle.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabStyle.Controls.Add($lblCurrentStyle)

$txtStyleSummary = New-Object System.Windows.Forms.TextBox
$txtStyleSummary.Location  = New-Object System.Drawing.Point(10, 242)
$txtStyleSummary.Size      = New-Object System.Drawing.Size(920, 90)
$txtStyleSummary.Multiline = $true
$txtStyleSummary.ReadOnly  = $true
$txtStyleSummary.ScrollBars= "Vertical"
$txtStyleSummary.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$txtStyleSummary.BackColor = [System.Drawing.Color]::FromArgb(245,248,255)
$txtStyleSummary.Text      = "(nessuno stile registrato ancora - invia qualche email modificata dopo Genera AI per iniziare)"
$tabStyle.Controls.Add($txtStyleSummary)
$script:txtStyleSummaryRef = $txtStyleSummary

$lblStyleHistory = New-Object System.Windows.Forms.Label
$lblStyleHistory.Text = "Cronologia evoluzione:"
$lblStyleHistory.Location = New-Object System.Drawing.Point(10, 346)
$lblStyleHistory.Size     = New-Object System.Drawing.Size(180, 26)
$lblStyleHistory.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabStyle.Controls.Add($lblStyleHistory)

$lstStyleHistory = New-Object System.Windows.Forms.ListBox
$lstStyleHistory.Location  = New-Object System.Drawing.Point(10, 376)
$lstStyleHistory.Size      = New-Object System.Drawing.Size(920, 260)
$lstStyleHistory.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lstStyleHistory.HorizontalScrollbar = $true
$tabStyle.Controls.Add($lstStyleHistory)
$script:lstStyleHistoryRef = $lstStyleHistory

$btnRefreshStyle = New-Btn "Aggiorna vista" 200 344 130 28 "SteelBlue"
$tabStyle.Controls.Add($btnRefreshStyle)

$btnEditStyle = New-Btn "Modifica manualmente" 340 344 160 28 "DarkSlateBlue"
$tabStyle.Controls.Add($btnEditStyle)

$btnClearStyle = New-Btn "Azzera profilo stile" 510 344 150 28 "Tomato"
$tabStyle.Controls.Add($btnClearStyle)

$btnEditStyle.Add_Click({
    $txtStyleSummary.ReadOnly = -not $txtStyleSummary.ReadOnly
    if ($txtStyleSummary.ReadOnly) {
        # Uscendo dalla modifica: salva
        $profile = Load-StyleProfile
        $profile.Summary = $txtStyleSummary.Text.Trim()
        $profile.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        Save-StyleProfile $profile
        $btnEditStyle.Text = "Modifica manualmente"
        $txtStyleSummary.BackColor = [System.Drawing.Color]::FromArgb(245,248,255)
        [System.Windows.Forms.MessageBox]::Show("Profilo di stile aggiornato manualmente.", "Salvato",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        $btnEditStyle.Text = "Salva modifica"
        $txtStyleSummary.BackColor = [System.Drawing.Color]::White
        $txtStyleSummary.Focus()
    }
})

$btnClearStyle.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Azzerare il profilo di stile? La cronologia rimane, ma il riassunto attuale sara' rimosso e ricomincera' da zero.",
        "Conferma azzeramento",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $profile = Load-StyleProfile
        $profile.Summary = ""
        $profile.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        Save-StyleProfile $profile
        Refresh-StyleTab
    }
})

function Refresh-StyleTab {
    $profile = Load-StyleProfile
    $lblStyleEmail.Text = "Profilo associato a: $($profile.Email)"
    $txtStyleSummary.Text = if ($profile.Summary -and $profile.Summary.Trim() -ne "") {
        $profile.Summary
    } else {
        "(nessuno stile registrato ancora - invia qualche email modificata dopo Genera AI per iniziare)"
    }
    $lstStyleHistory.Items.Clear()
    foreach ($h in ($profile.History | Sort-Object Date -Descending)) {
        $lstStyleHistory.Items.Add("$($h.Date)  -  $($h.Summary)") | Out-Null
    }
}

$btnRefreshStyle.Add_Click({ Refresh-StyleTab })

$lblApiKey = New-Label "API Key Claude:" 10 20 120 22
$tabSettings.Controls.Add($lblApiKey)
$txtApiKey = New-TextBox 140 17 600 22 $false $true
$tabSettings.Controls.Add($txtApiKey)

$btnSaveCreds = New-Btn "Salva Credenziali" 10 50 150 35 "SeaGreen"
$tabSettings.Controls.Add($btnSaveCreds)
$btnTestOutlook = New-Btn "Test Outlook" 170 50 130 35 "SteelBlue"
$tabSettings.Controls.Add($btnTestOutlook)

$lblOutlookStatus = New-Label "" 320 55 500 22 $false
$lblOutlookStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$tabSettings.Controls.Add($lblOutlookStatus)

# -- Credito e utilizzo API Claude --
$grpCredit = New-Object System.Windows.Forms.GroupBox
$grpCredit.Text     = "Credito e utilizzo API Claude"
$grpCredit.Location = New-Object System.Drawing.Point(10, 95)
$grpCredit.Size     = New-Object System.Drawing.Size(940, 110)
$tabSettings.Controls.Add($grpCredit)

$btnCheckCredit = New-Btn "Controlla utilizzo" 10 25 155 35 "DarkSlateBlue"
$grpCredit.Controls.Add($btnCheckCredit)

$lblBillingLink = New-Object System.Windows.Forms.LinkLabel
$lblBillingLink.Text      = "Apri console billing"
$lblBillingLink.Location  = New-Object System.Drawing.Point(175, 33)
$lblBillingLink.Size      = New-Object System.Drawing.Size(150, 20)
$lblBillingLink.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$lblBillingLink.Add_LinkClicked({ Start-Process "https://console.anthropic.com/settings/billing" })
$grpCredit.Controls.Add($lblBillingLink)

$lblAdminLink = New-Object System.Windows.Forms.LinkLabel
$lblAdminLink.Text      = "Chiave Admin (per utilizzo)"
$lblAdminLink.Location  = New-Object System.Drawing.Point(340, 33)
$lblAdminLink.Size      = New-Object System.Drawing.Size(200, 20)
$lblAdminLink.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAdminLink.Add_LinkClicked({ Start-Process "https://console.anthropic.com/settings/admin-keys" })
$grpCredit.Controls.Add($lblAdminLink)

$lblCreditVal = New-Object System.Windows.Forms.Label
$lblCreditVal.Location  = New-Object System.Drawing.Point(550, 30)
$lblCreditVal.Size      = New-Object System.Drawing.Size(375, 20)
$lblCreditVal.Text      = "(premi Controlla utilizzo per verificare la chiave)"
$lblCreditVal.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblCreditVal.ForeColor = [System.Drawing.Color]::Gray
$grpCredit.Controls.Add($lblCreditVal)
$script:lblCreditRef    = $lblCreditVal

$txtCredit = New-Object System.Windows.Forms.TextBox
$txtCredit.Location   = New-Object System.Drawing.Point(10, 60)
$txtCredit.Size       = New-Object System.Drawing.Size(916, 38)
$txtCredit.Multiline  = $true
$txtCredit.ReadOnly   = $true
$txtCredit.Font       = New-Object System.Drawing.Font("Consolas", 8)
$txtCredit.BackColor  = [System.Drawing.Color]::FromArgb(245, 248, 255)
$txtCredit.ScrollBars = "Vertical"
$txtCredit.Text       = "Nota: il dettaglio utilizzo richiede una chiave Admin (sk-ant-admin-...) dalla console Anthropic."
$grpCredit.Controls.Add($txtCredit)
$script:txtCreditRef  = $txtCredit

# -- Modello AI --
$grpModel = New-Object System.Windows.Forms.GroupBox
$grpModel.Text     = "Modello AI e preset generazione email"
$grpModel.Location = New-Object System.Drawing.Point(10, 215)
$grpModel.Size     = New-Object System.Drawing.Size(940, 130)
$tabSettings.Controls.Add($grpModel)

# Riga 1: Modello
$lblModel = New-Object System.Windows.Forms.Label
$lblModel.Text="Modello:"; $lblModel.Location=New-Object System.Drawing.Point(10,28)
$lblModel.Size=New-Object System.Drawing.Size(65,22); $lblModel.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpModel.Controls.Add($lblModel)

$cmbModel = New-Object System.Windows.Forms.ComboBox
$cmbModel.Location=New-Object System.Drawing.Point(80,25); $cmbModel.Size=New-Object System.Drawing.Size(380,24)
$cmbModel.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbModel.Font=New-Object System.Drawing.Font("Segoe UI",9)
@(
    "claude-haiku-4-5-20251001  |  Veloce ed economico  (~$0.25/M) -- consigliato per email",
    "claude-sonnet-4-6          |  Bilanciato           (~$3.00/M) -- qualita' superiore",
    "claude-opus-4-6            |  Massima qualita'     (~$15.0/M) -- testi complessi"
) | ForEach-Object { $cmbModel.Items.Add($_) | Out-Null }
$cmbModel.SelectedIndex = [int]$script:appSettings.Model
$grpModel.Controls.Add($cmbModel)
$script:cmbModelRef = $cmbModel

$lblModelHint = New-Object System.Windows.Forms.Label
$lblModelHint.Location=New-Object System.Drawing.Point(470,28); $lblModelHint.Size=New-Object System.Drawing.Size(310,22)
$lblModelHint.Text="Haiku = 12x piu' economico di Sonnet, ottimo per email"
$lblModelHint.Font=New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblModelHint.ForeColor=[System.Drawing.Color]::Gray
$grpModel.Controls.Add($lblModelHint)

# Riga 2: Lingua + Tono + Salva
$lblPLang = New-Object System.Windows.Forms.Label
$lblPLang.Text="Lingua default:"; $lblPLang.Location=New-Object System.Drawing.Point(10,62)
$lblPLang.Size=New-Object System.Drawing.Size(95,22); $lblPLang.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpModel.Controls.Add($lblPLang)

$cmbPresetLang = New-Object System.Windows.Forms.ComboBox
$cmbPresetLang.Location=New-Object System.Drawing.Point(110,59); $cmbPresetLang.Size=New-Object System.Drawing.Size(120,24)
$cmbPresetLang.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbPresetLang.Font=New-Object System.Drawing.Font("Segoe UI",9)
@("Italiano","English","Francais","Deutsch","Espanol") | ForEach-Object { $cmbPresetLang.Items.Add($_) | Out-Null }
$cmbPresetLang.SelectedItem = $script:appSettings.Language
if ($cmbPresetLang.SelectedIndex -lt 0) { $cmbPresetLang.SelectedIndex = 0 }
$grpModel.Controls.Add($cmbPresetLang)
$script:cmbPresetLangRef = $cmbPresetLang

$lblPTone = New-Object System.Windows.Forms.Label
$lblPTone.Text="Tono default:"; $lblPTone.Location=New-Object System.Drawing.Point(245,62)
$lblPTone.Size=New-Object System.Drawing.Size(85,22); $lblPTone.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpModel.Controls.Add($lblPTone)

$cmbPresetTone = New-Object System.Windows.Forms.ComboBox
$cmbPresetTone.Location=New-Object System.Drawing.Point(335,59); $cmbPresetTone.Size=New-Object System.Drawing.Size(130,24)
$cmbPresetTone.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbPresetTone.Font=New-Object System.Drawing.Font("Segoe UI",9)
@("Conciso","Professionale","Formale","Amichevole","Informale") | ForEach-Object { $cmbPresetTone.Items.Add($_) | Out-Null }
$cmbPresetTone.SelectedItem = $script:appSettings.Tone
if ($cmbPresetTone.SelectedIndex -lt 0) { $cmbPresetTone.SelectedIndex = 0 }
$grpModel.Controls.Add($cmbPresetTone)
$script:cmbPresetToneRef = $cmbPresetTone

$btnSavePreset = New-Object System.Windows.Forms.Button
$btnSavePreset.Text="Salva preset AI"; $btnSavePreset.Location=New-Object System.Drawing.Point(490,58)
$btnSavePreset.Size=New-Object System.Drawing.Size(140,28)
$btnSavePreset.BackColor=[System.Drawing.Color]::SeaGreen; $btnSavePreset.ForeColor=[System.Drawing.Color]::White
$btnSavePreset.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnSavePreset.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$grpModel.Controls.Add($btnSavePreset)

$lblPresetInfo = New-Object System.Windows.Forms.Label
$lblPresetInfo.Location=New-Object System.Drawing.Point(640,62); $lblPresetInfo.Size=New-Object System.Drawing.Size(280,22)
$lblPresetInfo.Font=New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblPresetInfo.ForeColor=[System.Drawing.Color]::Gray
$lblPresetInfo.Text="Questi valori appaiono pre-selezionati nel dialog AI"
$grpModel.Controls.Add($lblPresetInfo)

# Riga 3: Nome per firma
$lblSigName = New-Object System.Windows.Forms.Label
$lblSigName.Text="Nome per firma:"; $lblSigName.Location=New-Object System.Drawing.Point(10,96)
$lblSigName.Size=New-Object System.Drawing.Size(105,22); $lblSigName.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpModel.Controls.Add($lblSigName)

$txtSigName = New-Object System.Windows.Forms.TextBox
$txtSigName.Location=New-Object System.Drawing.Point(115,93); $txtSigName.Size=New-Object System.Drawing.Size(220,24)
$txtSigName.Font=New-Object System.Drawing.Font("Segoe UI",9)
$txtSigName.Text = if ($script:appSettings.ContainsKey('SignatureName')) { $script:appSettings.SignatureName } else { "" }
$grpModel.Controls.Add($txtSigName)
$script:txtSigNameRef = $txtSigName

$lblSigHint = New-Object System.Windows.Forms.Label
$lblSigHint.Location=New-Object System.Drawing.Point(345,96); $lblSigHint.Size=New-Object System.Drawing.Size(580,22)
$lblSigHint.Text="Se impostato, l'AI puo' firmare le email con questo nome. Se vuoto, non firma mai (a meno che tu non lo specifichi nel prompt)."
$lblSigHint.Font=New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblSigHint.ForeColor=[System.Drawing.Color]::Gray
$grpModel.Controls.Add($lblSigHint)

$txtSigName.Add_TextChanged({
    $script:appSettings.SignatureName = $txtSigName.Text.Trim()
    Save-AppSettings $script:appSettings
})

# -- Sincronizzazione posta --------------------------------------
$grpSync = New-Object System.Windows.Forms.GroupBox
$grpSync.Text     = "Sincronizzazione posta"
$grpSync.Location = New-Object System.Drawing.Point(10, 350)
$grpSync.Size     = New-Object System.Drawing.Size(940, 100)
$tabSettings.Controls.Add($grpSync)

$lblMaxEmails = New-Object System.Windows.Forms.Label
$lblMaxEmails.Text="Numero email da scaricare:"; $lblMaxEmails.Location=New-Object System.Drawing.Point(10,26)
$lblMaxEmails.Size=New-Object System.Drawing.Size(180,22); $lblMaxEmails.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpSync.Controls.Add($lblMaxEmails)

$numMaxEmails = New-Object System.Windows.Forms.NumericUpDown
$numMaxEmails.Location  = New-Object System.Drawing.Point(195, 23)
$numMaxEmails.Size      = New-Object System.Drawing.Size(70, 24)
$numMaxEmails.Minimum   = 5
$numMaxEmails.Maximum   = 500
$numMaxEmails.Increment = 5
$numMaxEmails.Value     = if ($script:appSettings.ContainsKey('MaxEmails')) { [int]$script:appSettings.MaxEmails } else { 30 }
$numMaxEmails.Font      = New-Object System.Drawing.Font("Segoe UI",9)
$grpSync.Controls.Add($numMaxEmails)
$script:numMaxEmailsRef = $numMaxEmails

$lblMaxEmailsHint = New-Object System.Windows.Forms.Label
$lblMaxEmailsHint.Location=New-Object System.Drawing.Point(280,26); $lblMaxEmailsHint.Size=New-Object System.Drawing.Size(500,22)
$lblMaxEmailsHint.Text="Applicato al prossimo avvio o al click su Aggiorna. Valori alti rallentano il caricamento."
$lblMaxEmailsHint.Font=New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblMaxEmailsHint.ForeColor=[System.Drawing.Color]::Gray
$grpSync.Controls.Add($lblMaxEmailsHint)

$numMaxEmails.Add_ValueChanged({
    $script:appSettings.MaxEmails = [int]$numMaxEmails.Value
    Save-AppSettings $script:appSettings
})

# -- Selezione account (supporta multi-account: O365, Gmail, IMAP, ecc.) --
$lblAccount = New-Object System.Windows.Forms.Label
$lblAccount.Text="Account da usare:"; $lblAccount.Location=New-Object System.Drawing.Point(10,60)
$lblAccount.Size=New-Object System.Drawing.Size(140,22); $lblAccount.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpSync.Controls.Add($lblAccount)

$cmbAccount = New-Object System.Windows.Forms.ComboBox
$cmbAccount.Location=New-Object System.Drawing.Point(155,57); $cmbAccount.Size=New-Object System.Drawing.Size(340,24)
$cmbAccount.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAccount.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpSync.Controls.Add($cmbAccount)
$script:cmbAccountRef = $cmbAccount

$btnRefreshAccounts = New-Object System.Windows.Forms.Button
$btnRefreshAccounts.Text="Rileva account"; $btnRefreshAccounts.Location=New-Object System.Drawing.Point(505,56)
$btnRefreshAccounts.Size=New-Object System.Drawing.Size(120,26)
$btnRefreshAccounts.BackColor=[System.Drawing.Color]::SteelBlue; $btnRefreshAccounts.ForeColor=[System.Drawing.Color]::White
$btnRefreshAccounts.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnRefreshAccounts.Font=New-Object System.Drawing.Font("Segoe UI",8)
$grpSync.Controls.Add($btnRefreshAccounts)

$lblAccountHint = New-Object System.Windows.Forms.Label
$lblAccountHint.Location=New-Object System.Drawing.Point(635,60); $lblAccountHint.Size=New-Object System.Drawing.Size(295,36)
$lblAccountHint.Text="Nessuna credenziale richiesta: usa gli account gia' configurati in Outlook (O365, Gmail, IMAP)."
$lblAccountHint.Font=New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblAccountHint.ForeColor=[System.Drawing.Color]::Gray
$grpSync.Controls.Add($lblAccountHint)

function Refresh-AccountList {
    $cmbAccount.Items.Clear()
    $cmbAccount.Items.Add("(Account predefinito)") | Out-Null
    $accs = Get-OutlookAccounts
    foreach ($a in $accs) {
        $label = if ($a.SmtpAddress) { "$($a.DisplayName)  <$($a.SmtpAddress)>" } else { $a.DisplayName }
        $cmbAccount.Items.Add($label) | Out-Null
    }
    $saved = if ($script:appSettings.ContainsKey('SelectedAccount')) { $script:appSettings.SelectedAccount } else { "" }
    $found = $false
    if ($saved -and $saved -ne "") {
        for ($i=0; $i -lt $cmbAccount.Items.Count; $i++) {
            if ($cmbAccount.Items[$i] -like "*$saved*") { $cmbAccount.SelectedIndex = $i; $found = $true; break }
        }
    }
    if (-not $found) { $cmbAccount.SelectedIndex = 0 }
}
Refresh-AccountList

$btnRefreshAccounts.Add_Click({ Refresh-AccountList })

$cmbAccount.Add_SelectedIndexChanged({
    if ($cmbAccount.SelectedIndex -eq 0) {
        $script:appSettings.SelectedAccount = ""
    } else {
        $accs = Get-OutlookAccounts
        $idx  = $cmbAccount.SelectedIndex - 1
        if ($idx -ge 0 -and $idx -lt $accs.Count) {
            $script:appSettings.SelectedAccount = if ($accs[$idx].SmtpAddress) { $accs[$idx].SmtpAddress } else { $accs[$idx].DisplayName }
        }
    }
    Save-AppSettings $script:appSettings
})

# -- Azzera dati personali ----------------------------------------
$grpReset = New-Object System.Windows.Forms.GroupBox
$grpReset.Text     = "Azzera dati personali"
$grpReset.Location = New-Object System.Drawing.Point(10, 460)
$grpReset.Size     = New-Object System.Drawing.Size(940, 50)
$tabSettings.Controls.Add($grpReset)

$lblResetInfo = New-Object System.Windows.Forms.Label
$lblResetInfo.Text      = "Rimuove: API Key, credenziali, cache contatti, bozze locali. Da usare prima di condividere lo script."
$lblResetInfo.Location  = New-Object System.Drawing.Point(10, 18)
$lblResetInfo.Size      = New-Object System.Drawing.Size(700, 20)
$lblResetInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblResetInfo.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
$grpReset.Controls.Add($lblResetInfo)

$btnResetData = New-Object System.Windows.Forms.Button
$btnResetData.Text      = "Azzera dati personali"
$btnResetData.Location  = New-Object System.Drawing.Point(720, 12)
$btnResetData.Size      = New-Object System.Drawing.Size(200, 28)
$btnResetData.BackColor = [System.Drawing.Color]::FromArgb(180,30,30)
$btnResetData.ForeColor = [System.Drawing.Color]::White
$btnResetData.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnResetData.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpReset.Controls.Add($btnResetData)

# -- Personalizzazione aspetto ----------------------------------------
$grpAppear = New-Object System.Windows.Forms.GroupBox
$grpAppear.Text     = "Personalizzazione aspetto"
$grpAppear.Location = New-Object System.Drawing.Point(10, 520)
$grpAppear.Size     = New-Object System.Drawing.Size(940, 130)
$tabSettings.Controls.Add($grpAppear)

# -- Riga 1: MOTD (testo header / form.Text) --
$lblMotd = New-Object System.Windows.Forms.Label
$lblMotd.Text="MOTD:"; $lblMotd.Location=New-Object System.Drawing.Point(10,22)
$lblMotd.Size=New-Object System.Drawing.Size(45,22); $lblMotd.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($lblMotd)

$txtAppName = New-Object System.Windows.Forms.TextBox
$txtAppName.Location  = New-Object System.Drawing.Point(58,19)
$txtAppName.Size      = New-Object System.Drawing.Size(540,24)
$txtAppName.Font      = New-Object System.Drawing.Font("Segoe UI",9)
$txtAppName.MaxLength = 255
$txtAppName.Text      = $script:appSettings.AppTitle
$grpAppear.Controls.Add($txtAppName)

# Limite caratteri live
$lblCharCount = New-Object System.Windows.Forms.Label
$lblCharCount.Location  = New-Object System.Drawing.Point(603,22)
$lblCharCount.Size      = New-Object System.Drawing.Size(60,20)
$lblCharCount.Font      = New-Object System.Drawing.Font("Segoe UI",8)
$lblCharCount.ForeColor = [System.Drawing.Color]::Gray
$lblCharCount.Text      = "$($txtAppName.Text.Length)/255"
$grpAppear.Controls.Add($lblCharCount)
$txtAppName.Add_TextChanged({
    $lblCharCount.Text = "$($txtAppName.Text.Length)/255"
    if ($txtAppName.Text.Length -gt 200) { $lblCharCount.ForeColor=[System.Drawing.Color]::Red }
    else { $lblCharCount.ForeColor=[System.Drawing.Color]::Gray }
})

# Font MOTD
$lblMFont = New-Object System.Windows.Forms.Label
$lblMFont.Text="Font:"; $lblMFont.Location=New-Object System.Drawing.Point(10,52)
$lblMFont.Size=New-Object System.Drawing.Size(38,22); $lblMFont.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($lblMFont)

$cmbMotdFont = New-Object System.Windows.Forms.ComboBox
$cmbMotdFont.Location=New-Object System.Drawing.Point(52,49); $cmbMotdFont.Size=New-Object System.Drawing.Size(160,24)
$cmbMotdFont.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbMotdFont.Font=New-Object System.Drawing.Font("Segoe UI",8)
@("Segoe UI","Aptos","Arial","Calibri","Tahoma","Verdana","Georgia","Courier New") |
    ForEach-Object { $cmbMotdFont.Items.Add($_) | Out-Null }
$savedMotdFont = if ($script:appSettings.MotdFont) { $script:appSettings.MotdFont } else { "Segoe UI" }
$fi = $cmbMotdFont.Items.IndexOf($savedMotdFont); if ($fi -ge 0) { $cmbMotdFont.SelectedIndex=$fi } else { $cmbMotdFont.SelectedIndex=0 }
$grpAppear.Controls.Add($cmbMotdFont)

# Size MOTD
$lblMSize = New-Object System.Windows.Forms.Label
$lblMSize.Text="Dim:"; $lblMSize.Location=New-Object System.Drawing.Point(222,52)
$lblMSize.Size=New-Object System.Drawing.Size(35,22); $lblMSize.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($lblMSize)

$cmbMotdSize = New-Object System.Windows.Forms.ComboBox
$cmbMotdSize.Location=New-Object System.Drawing.Point(260,49); $cmbMotdSize.Size=New-Object System.Drawing.Size(65,24)
$cmbMotdSize.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbMotdSize.Font=New-Object System.Drawing.Font("Segoe UI",8)
@("8","9","10","11","12","13","14","16","18","20","24","28") |
    ForEach-Object { $cmbMotdSize.Items.Add($_) | Out-Null }
$savedMotdSize = if ($script:appSettings.MotdSize) { $script:appSettings.MotdSize.ToString() } else { "14" }
$si = $cmbMotdSize.Items.IndexOf($savedMotdSize); if ($si -ge 0) { $cmbMotdSize.SelectedIndex=$si } else { $cmbMotdSize.SelectedIndex=5 }
$grpAppear.Controls.Add($cmbMotdSize)

# -- Riga 2: colore header + colore testo MOTD --
$lblHColor = New-Object System.Windows.Forms.Label
$lblHColor.Text="Colore header:"; $lblHColor.Location=New-Object System.Drawing.Point(10,85)
$lblHColor.Size=New-Object System.Drawing.Size(95,22); $lblHColor.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($lblHColor)

$btnPickHeader = New-Object System.Windows.Forms.Button
$btnPickHeader.Location=New-Object System.Drawing.Point(110,83); $btnPickHeader.Size=New-Object System.Drawing.Size(80,26)
$btnPickHeader.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnPickHeader.BackColor = Parse-Color $script:appSettings.HeaderColor
$btnPickHeader.Text="Scegli"
$grpAppear.Controls.Add($btnPickHeader)

$lblMotdColor = New-Object System.Windows.Forms.Label
$lblMotdColor.Text="Colore testo MOTD:"; $lblMotdColor.Location=New-Object System.Drawing.Point(205,85)
$lblMotdColor.Size=New-Object System.Drawing.Size(130,22); $lblMotdColor.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($lblMotdColor)

$btnPickMotdColor = New-Object System.Windows.Forms.Button
$savedMotdColor = if ($script:appSettings.MotdColor) { Parse-Color $script:appSettings.MotdColor } else { [System.Drawing.Color]::FromArgb(30,30,30) }
$btnPickMotdColor.Location=New-Object System.Drawing.Point(340,83); $btnPickMotdColor.Size=New-Object System.Drawing.Size(80,26)
$btnPickMotdColor.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnPickMotdColor.BackColor=$savedMotdColor; $btnPickMotdColor.Text="Scegli"
$grpAppear.Controls.Add($btnPickMotdColor)

# Pulsanti salva e reset
$btnSaveAppear = New-Object System.Windows.Forms.Button
$btnSaveAppear.Text="Salva aspetto"; $btnSaveAppear.Location=New-Object System.Drawing.Point(640,83)
$btnSaveAppear.Size=New-Object System.Drawing.Size(130,26)
$btnSaveAppear.BackColor=[System.Drawing.Color]::SteelBlue; $btnSaveAppear.ForeColor=[System.Drawing.Color]::White
$btnSaveAppear.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnSaveAppear.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$grpAppear.Controls.Add($btnSaveAppear)

$btnResetAppear = New-Object System.Windows.Forms.Button
$btnResetAppear.Text="Reset default"; $btnResetAppear.Location=New-Object System.Drawing.Point(780,83)
$btnResetAppear.Size=New-Object System.Drawing.Size(130,26)
$btnResetAppear.BackColor=[System.Drawing.Color]::DarkGray; $btnResetAppear.ForeColor=[System.Drawing.Color]::White
$btnResetAppear.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnResetAppear.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grpAppear.Controls.Add($btnResetAppear)

# Riferimenti script scope per handler
$script:cmbMotdFontRef  = $cmbMotdFont
$script:cmbMotdSizeRef  = $cmbMotdSize
$script:btnPickMotdRef  = $btnPickMotdColor

$grpChangelog = New-Object System.Windows.Forms.GroupBox
$grpChangelog.Text = "Changelog Versioni"
$grpChangelog.Location = New-Object System.Drawing.Point(10, 628)
$grpChangelog.Size = New-Object System.Drawing.Size(940, 200)

$txtChangelog = New-Object System.Windows.Forms.TextBox
$txtChangelog.Multiline = $true
$txtChangelog.ScrollBars = "Vertical"
$txtChangelog.ReadOnly = $true
$txtChangelog.Font = New-Object System.Drawing.Font("Consolas", 8)
$txtChangelog.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)
$txtChangelog.Location = New-Object System.Drawing.Point(10, 20)
$txtChangelog.Size = New-Object System.Drawing.Size(918, 228)

# Leggi changelog da file esterno o usa quello integrato
$changelogFile = "$rootPath\changelog.txt"
if (Test-Path $changelogFile) {
    $txtChangelog.Text = [System.IO.File]::ReadAllText($changelogFile, [System.Text.Encoding]::UTF8)
} else {
    $txtChangelog.Text = @"
==============================================================
   LIGHT MAIL CLIENT + AI - CHANGELOG VERSIONI
==============================================================

v20 (2026-09)
  [+] Pannello "Credito e utilizzo API Claude" nel tab Impostazioni
  [+] Pulsante Controlla utilizzo (Admin Key) + link billing diretto
  [+] Selezione modello AI: Haiku / Sonnet / Opus con prezzi indicativi
  [*] Default modello cambiato a Haiku (12x piu' economico per email)
  [*] Fix encoding UTF-8 BOM per compatibilita' PowerShell 5
  [*] Fix em dash e caratteri speciali che causavano errori di parsing

v19 (2026-09)
  [+] Dialog AI con selezione lingua (IT/EN/FR/DE/ES) e tono
  [*] Lingua rispettata correttamente (system prompt in inglese neutro)
  [*] Encoding HTTP fix: HttpWebRequest UTF-8 senza BOM (era WebClient UTF-16)
  [*] Fix corpo email vuoto in invio risposta (closure scope txtRBody)
  [*] Fix originalEmail closure scope (script:rOrigEmail)
  [*] Reply: To/CC/Subject popolati correttamente
  [*] ReplyAll: To include mittente + destinatari originali
  [*] Email originale in corpo risposta con header stile Outlook

v15 (2026-09)
  [+] Posta in Arrivo: ListView a 4 colonne
      Mittente | Oggetto | Data | Stato
  [+] Ordinamento nativo al click sull'header colonna
  [+] Email non lette in blu + grassetto
  [+] Griglia con linee e full row select
  [+] Tab Impostazioni: area Changelog Versioni

v14m (2026-08)
  [*] Fix gestione immagini CID e firma HTML
  [*] Bozze JSON con aggiornamento live della lista
  [+] Pulsante Carica Bozza nel tab Bozze
  [*] Fix scope closure variabili Reply
  [*] Vari fix stabilita' e compatibilita'

v14 (2026-07)
  [+] Sistema Bozze ibrido: JSON locale + cartella Bozze Outlook
  [+] Risoluzione Exchange Distinguished Name mittente
  [*] Migliorata gestione allegati multipli

v13 (2026-06)
  [+] Integrazione AI per risposte automatiche
  [+] Tab Reply dinamico per risposte AI
  [+] Configurazione API Key nel tab Impostazioni

v12 (2026-05)
  [+] Salvataggio bozze nella cartella Bozze di Outlook
  [+] Cifratura credenziali con DPAPI Windows
  [*] Migliorata gestione sessione COM Outlook

v11 (2026-04)
  [+] Lettura email dalla cartella Inbox
  [+] Anteprima corpo email HTML con WebBrowser
  [+] Visualizzazione mittente e data ricezione

v10 (2026-03)
  [+] Prima versione funzionante
  [+] Compose email con campi To, CC, Subject, Body
  [+] Invio email con allegati tramite Outlook COM
"@
}

$grpChangelog.Controls.Add($txtChangelog)
$tabSettings.Controls.Add($grpChangelog)
# ============================================================
# DIALOG AI RIUTILIZZABILE (Compose + Reply)
# ============================================================
function Show-AIDialog {
    param($title = "Genera con AI", $defaultPrompt = "")

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = $title
    $dlg.Size            = New-Object System.Drawing.Size(520, 310)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false

    # Prompt
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text     = "Istruzioni:"
    $lbl1.Location = New-Object System.Drawing.Point(10, 10)
    $lbl1.Size     = New-Object System.Drawing.Size(490, 18)
    $dlg.Controls.Add($lbl1)

    $txtP = New-Object System.Windows.Forms.TextBox
    $txtP.Location   = New-Object System.Drawing.Point(10, 30)
    $txtP.Size       = New-Object System.Drawing.Size(490, 60)
    $txtP.Multiline  = $true
    $txtP.ScrollBars = "Vertical"
    $txtP.Text       = $defaultPrompt
    $dlg.Controls.Add($txtP)

    # Lingua
    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text     = "Lingua:"
    $lbl2.Location = New-Object System.Drawing.Point(10, 102)
    $lbl2.Size     = New-Object System.Drawing.Size(60, 20)
    $dlg.Controls.Add($lbl2)

    $cmbLang = New-Object System.Windows.Forms.ComboBox
    $cmbLang.Location     = New-Object System.Drawing.Point(75, 99)
    $cmbLang.Size         = New-Object System.Drawing.Size(140, 24)
    $cmbLang.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("Italiano","English","Francais","Deutsch","Espanol") | ForEach-Object { $cmbLang.Items.Add($_) | Out-Null }
    $cmbLang.SelectedIndex = 0
    $presetLang = if ($script:appSettings) { $script:appSettings.Language } else { "Italiano" }
    $li = $cmbLang.FindStringExact($presetLang)
    if ($li -ge 0) { $cmbLang.SelectedIndex = $li }
    $dlg.Controls.Add($cmbLang)

    # Tono
    $lbl3 = New-Object System.Windows.Forms.Label
    $lbl3.Text     = "Tono:"
    $lbl3.Location = New-Object System.Drawing.Point(240, 102)
    $lbl3.Size     = New-Object System.Drawing.Size(45, 20)
    $dlg.Controls.Add($lbl3)

    $cmbTone = New-Object System.Windows.Forms.ComboBox
    $cmbTone.Location      = New-Object System.Drawing.Point(290, 99)
    $cmbTone.Size          = New-Object System.Drawing.Size(210, 24)
    $cmbTone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("Professionale","Formale","Amichevole","Informale","Conciso") | ForEach-Object { $cmbTone.Items.Add($_) | Out-Null }
    $cmbTone.SelectedIndex = 0
    $presetTone = if ($script:appSettings) { $script:appSettings.Tone } else { "Conciso" }
    $ti = $cmbTone.FindStringExact($presetTone)
    if ($ti -ge 0) { $cmbTone.SelectedIndex = $ti }
    $dlg.Controls.Add($cmbTone)

    # Anteprima tono
    $lblToneHint = New-Object System.Windows.Forms.Label
    $lblToneHint.Location  = New-Object System.Drawing.Point(10, 132)
    $lblToneHint.Size      = New-Object System.Drawing.Size(490, 18)
    $lblToneHint.ForeColor = [System.Drawing.Color]::Gray
    $lblToneHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblToneHint.Text      = "Cordiale e chiaro, adatto alla maggior parte delle email aziendali."
    $dlg.Controls.Add($lblToneHint)

    $toneHints = @{
        "Professionale" = "Cordiale e chiaro, adatto alla maggior parte delle email aziendali."
        "Formale"       = "Distaccato e formale, per contesti istituzionali o legali."
        "Amichevole"    = "Caldo e diretto, meno rigido del professionale."
        "Informale"     = "Colloquiale, come tra colleghi che si conoscono bene."
        "Conciso"       = "Brevissimo: 2-4 frasi, niente fronzoli."
    }
    $cmbTone.Add_SelectedIndexChanged({
        $lblToneHint.Text = $toneHints[$cmbTone.SelectedItem.ToString()]
    })

    # Separatore
    $sep = New-Object System.Windows.Forms.Label
    $sep.Location  = New-Object System.Drawing.Point(10, 158)
    $sep.Size      = New-Object System.Drawing.Size(490, 1)
    $sep.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dlg.Controls.Add($sep)

    # Bottoni
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Genera"
    $btnOK.Location     = New-Object System.Drawing.Point(290, 170)
    $btnOK.Size         = New-Object System.Drawing.Size(100, 32)
    $btnOK.BackColor    = [System.Drawing.Color]::MediumPurple
    $btnOK.ForeColor    = [System.Drawing.Color]::White
    $btnOK.FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annulla"
    $btnCancel.Location     = New-Object System.Drawing.Point(400, 170)
    $btnCancel.Size         = New-Object System.Drawing.Size(100, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel

    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            Prompt   = $txtP.Text.Trim()
            Language = $cmbLang.SelectedItem.ToString()
            Tone     = $cmbTone.SelectedItem.ToString()
        }
    }
    return $null
}

# ============================================================
# EVENT HANDLERS - TAB COMPOSE
# ============================================================
$btnAddAtt.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq "OK") {
        foreach ($f in $dlg.FileNames) {
            $script:attachPaths.Add($f) | Out-Null
            $lstAttach.Items.Add([System.IO.Path]::GetFileName($f)) | Out-Null
        }
    }
})

$btnRemAtt.Add_Click({
    if ($lstAttach.SelectedIndex -ge 0) {
        $idx = $lstAttach.SelectedIndex
        $script:attachPaths.RemoveAt($idx)
        $lstAttach.Items.RemoveAt($idx)
    }
})

$btnAI.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:config.ApiKey)) {
        [System.Windows.Forms.MessageBox]::Show("Configura la API Key nel tab Impostazioni.", "API Key mancante",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $opts = Show-AIDialog -title "Genera email con AI"
    if ($null -eq $opts) { return }
    $p = if ([string]::IsNullOrWhiteSpace($opts.Prompt)) { "scrivi una email" } else { $opts.Prompt }

    # Leggi il testo incollato nel corpo (eventuale email originale)
    $existingText = Get-ComposePlain
    $existingHtml = Get-ComposeHtml

    try {
        if (-not [string]::IsNullOrWhiteSpace($existingText) -and $existingText.Trim() -ne "") {
            # C'e' testo nel corpo -> trattalo come email originale, genera risposta sopra
            $replyContext = $existingText.Trim()
            $generated = Invoke-AI -prompt $p -apiKey $script:config.ApiKey `
                                   -replyContext $replyContext `
                                   -language $opts.Language -tone $opts.Tone `
                                   -mode "reply"

            # Separa risposta + testo originale in HTML
            $generatedHtml = $generated -replace "`r`n","<br>" -replace "`n","<br>"
            $separator = "<hr style='border:none;border-top:1px solid #ccc;margin:12px 0;'>"
            $quotedHtml = "<div style='color:#666;font-size:9pt;'>$existingHtml</div>"
            Set-ComposeHtml "$generatedHtml$separator$quotedHtml"
            $script:lastAIGenerated = $generated.Trim()
        } else {
            # Corpo vuoto -> genera email nuova
            $generated = Invoke-AI -prompt $p -apiKey $script:config.ApiKey `
                                   -language $opts.Language -tone $opts.Tone `
                                   -mode "new"
            $generatedHtml = $generated -replace "`r`n","<br>" -replace "`n","<br>"
            Set-ComposeHtml $generatedHtml
            $script:lastAIGenerated = $generated.Trim()
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore AI:`n$_", "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnSend.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtTo.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Inserisci almeno un destinatario.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    if ([string]::IsNullOrWhiteSpace($txtSubj.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Inserisci un oggetto.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $finalPlain = Get-ComposePlain
    if ([string]::IsNullOrWhiteSpace($finalPlain)) {
        [System.Windows.Forms.MessageBox]::Show("Il corpo del messaggio e' vuoto.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    $htmlBody = Get-ComposeHtml

    # Apprendimento stile: confronta versione AI vs versione finale modificata dall'utente
    if ($script:appSettings.StyleLearningEnabled -and $script:lastAIGenerated -and -not [string]::IsNullOrWhiteSpace($script:lastAIGenerated)) {
        try { Update-StyleProfile -aiText $script:lastAIGenerated -userText $finalPlain -apiKey $script:config.ApiKey } catch {}
        $script:lastAIGenerated = $null
    }

    if (Send-OutlookEmail -to $txtTo.Text -cc $txtCc.Text -subject $txtSubj.Text -body $htmlBody -attachments $script:attachPaths) {
        $txtTo.Text=""; $txtCc.Text=""; $txtSubj.Text=""
        Set-ComposeHtml ""
        $script:attachPaths.Clear(); $lstAttach.Items.Clear()
    }
})

$btnSaveDraft.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSubj.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Inserisci un oggetto per la bozza.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return
    }
    Save-Draft -to $txtTo.Text -cc $txtCc.Text -subject $txtSubj.Text -body (Get-ComposeHtml)
    Refresh-DraftsList
    [System.Windows.Forms.MessageBox]::Show("Bozza salvata nel tab Bozze!", "Successo",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnClear.Add_Click({
    # Pulisci campi email
    $txtTo.Text=""; $txtCc.Text=""; $txtSubj.Text=""
    $script:attachPaths.Clear(); $lstAttach.Items.Clear()

    # Reset toolbar ai valori default
    $cmbFont.SelectedIndex = 0          # Aptos
    $cmbSize.SelectedIndex = 4          # 12pt
    $cmbLS.SelectedIndex   = 4          # 1.15
    $btnFgColor.ForeColor  = [System.Drawing.Color]::Red
    $btnBgColor.BackColor  = [System.Drawing.Color]::Yellow
    $script:fgColor = "#000000"
    $script:bgColor = "#FFFF00"

    # Ricarica il WebBrowser con HTML pulito e stili default
    $script:webCompose.DocumentText = @"
<html>
<head>
<style>
  body {
    font-family: Aptos, 'Segoe UI', sans-serif;
    font-size: 12pt;
    line-height: 1.15;
    margin: 8px;
    color: #222;
    outline: none;
  }
</style>
</head>
<body contenteditable="true" id="mailbody"><br></body>
</html>
"@
})

# ============================================================
# EVENT HANDLERS - TAB INBOX
# ============================================================
$btnRefresh.Add_Click({
    $btnRefresh.Enabled = $false
    $script:progInboxRef.Value   = 0
    $script:progInboxRef.Visible = $true
    $lblStatus.Text = "Aggiornamento in corso..."
    $script:progInboxRef.Refresh()

    $inboxProgress = {
        param($current, $total)
        if ($total -gt 0) {
            $pct = [int](($current / $total) * 100)
            $script:progInboxRef.Value = [Math]::Min($pct, 100)
            $lblStatus.Text = "Caricamento email $current di $total..."
            $script:progInboxRef.Refresh()
            $lblStatus.Refresh()
        }
    }

    try {
        Invoke-InboxRefresh -progressCallback $inboxProgress
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore durante l'aggiornamento: $_", "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    $script:progInboxRef.Visible = $false
    $btnRefresh.Enabled = $true
})

$lstEmails.Add_SelectedIndexChanged({
    $idx = if ($lstEmails.SelectedItems.Count -gt 0) {
        $lstEmails.SelectedItems[0].Index
    } else { -1 }
    if ($idx -ge 0 -and $idx -lt $script:inboxEmails.Count) {
        $e = $script:inboxEmails[$idx]
        $script:selectedEmail = $e

        $txtReadSubj.Text = $e.Subject
        $txtReadFrom.Text = "$($e.FromName)  <$($e.From)>"

        # FIX: WebBrowser refresh con Timer + blocco immagini remote di default
        if ($e.HTMLBody -and $e.HTMLBody.Trim() -ne "") {
            $script:currentEmailHtmlOrig = $e.HTMLBody
            $blockedHtml = Block-RemoteImages $e.HTMLBody
            $script:pendingHTML = $blockedHtml
            # Mostra il pulsante solo se ci sono davvero immagini remote da bloccare
            if ($blockedHtml -ne $e.HTMLBody) {
                $script:btnLoadImagesRef.Visible = $true
            } else {
                $script:btnLoadImagesRef.Visible = $false
            }
        } else {
            $script:currentEmailHtmlOrig = ""
            $script:btnLoadImagesRef.Visible = $false
            $plain = $e.Body -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace "`n","<br>"
            $script:pendingHTML = @"
<html><head><style>body{font-family:'Segoe UI';font-size:10pt;margin:15px;color:#333;}</style></head>
<body>$plain</body></html>
"@
        }
        $script:webTimer.Stop()
        $webBody.DocumentText = "<html><body></body></html>"
        $script:webTimer.Start()

        # Allegati
        $lstAttachRead.Items.Clear()
        if ($e.Attachments -and $e.Attachments.Count -gt 0) {
            foreach ($att in $e.Attachments) {
                $sizeKB = [Math]::Round($att.Size / 1024, 1)
                $lstAttachRead.Items.Add("$($att.Name)  ($sizeKB KB)") | Out-Null
            }
        } else {
            $lstAttachRead.Items.Add("(nessun allegato)") | Out-Null
        }

        # Abilita bottoni Reply
        $btnReply.Enabled = $true
        $btnReplyAll.Enabled = $true
        $btnForward.Enabled = $true
    }
})

# FIX v13: Reply apre TAB DEDICATO, non "Nuova Email"
$btnReply.Add_Click({
    if ($null -eq $script:selectedEmail) {
        [System.Windows.Forms.MessageBox]::Show("Seleziona prima un'email dalla lista.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Open-ReplyTab -originalEmail $script:selectedEmail -replyAll $false -forward $false
})

$btnReplyAll.Add_Click({
    if ($null -eq $script:selectedEmail) { return }
    Open-ReplyTab -originalEmail $script:selectedEmail -replyAll $true -forward $false
})

$btnForward.Add_Click({
    if ($null -eq $script:selectedEmail) { return }
    Open-ReplyTab -originalEmail $script:selectedEmail -replyAll $false -forward $true
})

$btnSaveAttach.Add_Click({
    $idx = $lstAttachRead.SelectedIndex
    if ($idx -lt 0 -or $null -eq $script:selectedEmail) {
        [System.Windows.Forms.MessageBox]::Show("Seleziona un allegato dalla lista.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    try {
        $e = $script:selectedEmail
        $att = $e.Attachments[$idx]
        $outlook = Get-OutlookInstance
        $namespace = $outlook.GetNamespace("MAPI")
        $mail = $namespace.GetItemFromID($e.Id)
        
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.FileName = $att.Name
        if ($dlg.ShowDialog() -eq "OK") {
            $mail.Attachments.Item($att.Index).SaveAsFile($dlg.FileName)
            [System.Windows.Forms.MessageBox]::Show("Allegato salvato!", "Successo",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore salvataggio:`n$_", "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ============================================================
# EVENT HANDLERS - TAB DRAFTS
# ============================================================
$btnLoadDraft.Add_Click({
    $idx = $script:lstDraftsRef.SelectedIndex
    if ($idx -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Seleziona una bozza dalla lista.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # Salta la riga placeholder "(nessuna bozza salvata)"
    $selText = $script:lstDraftsRef.Items[$idx].ToString()
    if ($selText.StartsWith("(")) { return }

    $drafts = Load-Drafts
    if ($drafts.Count -eq 0 -or $idx -ge $drafts.Count) {
        [System.Windows.Forms.MessageBox]::Show("Bozza non trovata. Prova ad aggiornare la lista.", "Errore",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $draft = $drafts[$idx]
    $tab.SelectedTab    = $tabCompose
    $txtTo.Text         = if ($draft.To)      { $draft.To }     else { "" }
    $txtCc.Text         = if ($draft.Cc)      { $draft.Cc }     else { "" }
    $txtSubj.Text       = if ($draft.Subject) { $draft.Subject } else { "" }
    Set-ComposeHtml     (if ($draft.Body)    { $draft.Body }    else { "" })
    $script:currentDraftId = $draft.Id
    [System.Windows.Forms.MessageBox]::Show("Bozza caricata nel tab Nuova Email.", "Caricata",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnDeleteDraft.Add_Click({
    $idx = $script:lstDraftsRef.SelectedIndex
    if ($idx -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Seleziona una bozza da eliminare.", "Attenzione",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selText = $script:lstDraftsRef.Items[$idx].ToString()
    if ($selText.StartsWith("(")) { return }

    $drafts = Load-Drafts
    if ($drafts.Count -eq 0 -or $idx -ge $drafts.Count) { return }
    $draft = $drafts[$idx]
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Eliminare la bozza:`n$($draft.Subject)?",
        "Conferma eliminazione",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        Delete-Draft -id $draft.Id
        Refresh-DraftsList
    }
})

$btnResetDrafts.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Cancellare TUTTE le bozze salvate?`nQuesta operazione non e' reversibile.",
        "Conferma reset",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        if (Test-Path $draftsFile) { Remove-Item $draftsFile -Force }
        Refresh-DraftsList
        [System.Windows.Forms.MessageBox]::Show("Lista bozze azzerata.", "Fatto",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# ============================================================
# EVENT HANDLERS - TAB SETTINGS
# ============================================================
# -- Handler: Salva preset AI ---------------------------------
$btnSavePreset.Add_Click({
    $script:appSettings.Model    = $script:cmbModelRef.SelectedIndex
    $script:appSettings.Language = $script:cmbPresetLangRef.SelectedItem.ToString()
    $script:appSettings.Tone     = $script:cmbPresetToneRef.SelectedItem.ToString()
    Save-AppSettings $script:appSettings
    [System.Windows.Forms.MessageBox]::Show(
        "Preset AI salvato:`nModello: $($script:cmbModelRef.SelectedItem.ToString().Split('|')[0].Trim())`nLingua: $($script:appSettings.Language)`nTono: $($script:appSettings.Tone)",
        "Preset salvato", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# -- Handler: colore header ------------------------------------
$btnPickHeader.Add_Click({
    $d = New-Object System.Windows.Forms.ColorDialog
    $d.Color = $pnlHeader.BackColor
    if ($d.ShowDialog() -eq "OK") {
        $pnlHeader.BackColor     = $d.Color
        $btnPickHeader.BackColor = $d.Color
    }
})

$btnPickMotdColor.Add_Click({
    $d = New-Object System.Windows.Forms.ColorDialog
    $d.Color = $lblAppTitle.ForeColor
    if ($d.ShowDialog() -eq "OK") {
        $lblAppTitle.ForeColor           = $d.Color
        $script:btnPickMotdRef.BackColor = $d.Color
    }
})

$btnSaveAppear.Add_Click({
    $hc    = $pnlHeader.BackColor
    $mc    = $lblAppTitle.ForeColor
    $mFont = $script:cmbMotdFontRef.SelectedItem.ToString()
    $mSize = [int]$script:cmbMotdSizeRef.SelectedItem.ToString()
    $title = if ($txtAppName.Text.Trim()) { $txtAppName.Text.Trim() } else { "Mail AI Assistant" }
    $script:appSettings.HeaderColor = "$($hc.R),$($hc.G),$($hc.B)"
    $script:appSettings.MotdColor   = "$($mc.R),$($mc.G),$($mc.B)"
    $script:appSettings.AppTitle    = $title
    $script:appSettings.FormTitle   = $title
    $script:appSettings.MotdFont    = $mFont
    $script:appSettings.MotdSize    = $mSize
    $lblAppTitle.Text      = $title
    $lblAppTitle.ForeColor = $mc
    $lblAppTitle.Font      = New-Object System.Drawing.Font($mFont, $mSize, [System.Drawing.FontStyle]::Bold)
    $form.Text             = $title
    Save-AppSettings $script:appSettings
    [System.Windows.Forms.MessageBox]::Show("Aspetto salvato e applicato.", "Salvato",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnResetAppear.Add_Click({
    $defHeader    = [System.Drawing.Color]::FromArgb(230, 0, 0)
    $defMotdColor = [System.Drawing.Color]::White
    $defTitle     = "Mail AI Assistant"
    $defFont      = "Segoe UI"
    $defSize      = 14
    $pnlHeader.BackColor             = $defHeader
    $btnPickHeader.BackColor         = $defHeader
    $script:btnPickMotdRef.BackColor = $defMotdColor
    $txtAppName.Text                 = $defTitle
    $lblAppTitle.Text                = $defTitle
    $lblAppTitle.ForeColor           = $defMotdColor
    $lblAppTitle.Font                = New-Object System.Drawing.Font($defFont, $defSize, [System.Drawing.FontStyle]::Bold)
    $form.Text                       = $defTitle
    $fi = $script:cmbMotdFontRef.Items.IndexOf($defFont)
    if ($fi -ge 0) { $script:cmbMotdFontRef.SelectedIndex = $fi }
    $si = $script:cmbMotdSizeRef.Items.IndexOf($defSize.ToString())
    if ($si -ge 0) { $script:cmbMotdSizeRef.SelectedIndex = $si }
    $script:appSettings.HeaderColor = "230,0,0"
    $script:appSettings.MotdColor   = "255,255,255"
    $script:appSettings.AppTitle    = $defTitle
    $script:appSettings.FormTitle   = $defTitle
    $script:appSettings.MotdFont    = $defFont
    $script:appSettings.MotdSize    = $defSize
    Save-AppSettings $script:appSettings
    [System.Windows.Forms.MessageBox]::Show("Aspetto ripristinato ai valori default.", "Reset",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})


$btnSaveCreds.Add_Click({
    $apiKeyValue = $txtApiKey.Text.Trim()
    Save-Credentials -apiKey $apiKeyValue
    # Aggiorna script:config in memoria subito!
    $script:config = @{ ApiKey = $apiKeyValue }
    [System.Windows.Forms.MessageBox]::Show(
        "Credenziali salvate!`nAPI Key aggiornata in memoria.",
        "Successo",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnTestOutlook.Add_Click({
    try {
        $outlook = Get-OutlookInstance
        if ($outlook) {
            $namespace = $outlook.GetNamespace("MAPI")
            $inbox = $namespace.GetDefaultFolder(6)
            $lblOutlookStatus.Text = "OK - $($inbox.Items.Count) email in inbox"
            $lblOutlookStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblOutlookStatus.Text = "ERRORE - Outlook non disponibile"
            $lblOutlookStatus.ForeColor = [System.Drawing.Color]::Red
        }
    } catch {
        $lblOutlookStatus.Text = "ERRORE - $_"
        $lblOutlookStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnCheckCredit.Add_Click({
    $apiKey = $txtApiKey.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        $apiKey = if ($script:config) { $script:config.ApiKey } else { "" }
    }
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Inserisci la API Key nel campo sopra.",
            "API Key mancante",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $script:lblCreditRef.Text      = "Verifica in corso..."
    $script:lblCreditRef.ForeColor = [System.Drawing.Color]::Gray
    $script:txtCreditRef.Text      = "Contatto api.anthropic.com..."
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false

    try {
        # Verifica chiave con chiamata minima (1 token Haiku = ~$0.000000025)
        $tp     = '{"model":"claude-haiku-4-5-20251001","max_tokens":1,"messages":[{"role":"user","content":"hi"}]}'
        $tbytes = $utf8NoBom.GetBytes($tp)
        $req = [System.Net.HttpWebRequest]::Create("https://api.anthropic.com/v1/messages")
        $req.Method = "POST"; $req.ContentType = "application/json; charset=utf-8"
        $req.ContentLength = $tbytes.Length; $req.Timeout = 10000
        $req.Headers.Add("x-api-key",         $apiKey)
        $req.Headers.Add("anthropic-version", "2023-06-01")
        $s = $req.GetRequestStream(); $s.Write($tbytes,0,$tbytes.Length); $s.Close()
        $req.GetResponse().Close()

        $script:lblCreditRef.Text      = "API Key valida"
        $script:lblCreditRef.ForeColor = [System.Drawing.Color]::DarkGreen
        $script:lblCreditRef.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

        # Prova a leggere utilizzo -- richiede Admin Key (sk-ant-admin-...)
        $isAdmin = $apiKey.StartsWith("sk-ant-admin")
        if ($isAdmin) {
            try {
                $now       = Get-Date
                $startDate = (Get-Date -Year $now.Year -Month $now.Month -Day 1).ToString("yyyy-MM-dd")
                $endDate   = $now.ToString("yyyy-MM-dd")
                $r = [System.Net.HttpWebRequest]::Create(
                    "https://api.anthropic.com/v1/usage?start_date=$startDate&end_date=$endDate")
                $r.Method = "GET"; $r.Timeout = 10000
                $r.Headers.Add("x-api-key",         $apiKey)
                $r.Headers.Add("anthropic-version", "2023-06-01")
                $resp   = $r.GetResponse()
                $reader = New-Object System.IO.StreamReader($resp.GetResponseStream(), $utf8NoBom)
                $usage  = $reader.ReadToEnd() | ConvertFrom-Json
                $reader.Close(); $resp.Close()

                $pricing = @{
                    "claude-sonnet-4-6"         = @{ i=3.00;  o=15.00 }
                    "claude-haiku-4-5-20251001" = @{ i=0.25;  o=1.25  }
                    "claude-opus-4-6"           = @{ i=15.00; o=75.00 }
                }
                $totalCost = 0.0
                $out = "Utilizzo $($now.ToString('MMMM yyyy')):`r`n"
                if ($usage.data -and $usage.data.Count -gt 0) {
                    $byM = @{}
                    foreach ($e in $usage.data) {
                        if (-not $byM[$e.model]) { $byM[$e.model]=@{i=0L;o=0L} }
                        $byM[$e.model].i += if ($e.input_tokens)  { [long]$e.input_tokens }  else { 0L }
                        $byM[$e.model].o += if ($e.output_tokens) { [long]$e.output_tokens } else { 0L }
                    }
                    foreach ($mid in ($byM.Keys | Sort-Object)) {
                        $mv = $byM[$mid]
                        $p  = if ($pricing[$mid]) { $pricing[$mid] } else { @{i=3.0;o=15.0} }
                        $ct = ($mv.i/1000000.0)*$p.i + ($mv.o/1000000.0)*$p.o
                        $totalCost += $ct
                        $sn = $mid -replace 'claude-','' -replace '-\d{8,}',''
                        $out += ("  {0,-24} In:{1,8:N0}  Out:{2,7:N0}  USD:{3:N4}`r`n" -f $sn,$mv.i,$mv.o,$ct)
                    }
                    $out += ("  Totale: `${0:N4}  (~EUR {1:N4})" -f $totalCost,($totalCost*0.92))
                } else {
                    $out += "  Nessuna chiamata API nel mese corrente."
                }
                $script:txtCreditRef.Text = $out
            } catch {
                $script:txtCreditRef.Text = "Errore lettura utilizzo: $($_.Exception.Message)"
            }
        } else {
            $script:txtCreditRef.Text = "Chiave API standard -- statistiche non disponibili.`r`n" +
                "Per vedere l'utilizzo: crea una Admin Key (sk-ant-admin-...) su console.anthropic.com/settings/admin-keys`r`n" +
                "oppure consulta direttamente console.anthropic.com/settings/billing"
        }

    } catch [System.Net.WebException] {
        $ec = 0; $eb = ""
        try {
            $ec = [int]$_.Exception.Response.StatusCode
            $es = $_.Exception.Response.GetResponseStream()
            $er = New-Object System.IO.StreamReader($es, $utf8NoBom)
            $eb = $er.ReadToEnd(); $er.Close()
        } catch { }
        $msg = switch ($ec) {
            401 { "API Key non valida (401). Controlla la chiave su console.anthropic.com." }
            403 { "Accesso negato (403). Chiave disabilitata o senza permessi." }
            429 { "Rate limit (429). Attendi qualche secondo e riprova." }
            0   { "Errore di rete: $($_.Exception.Message)" }
            default { $d=""; try{$d=($eb|ConvertFrom-Json).error.message}catch{}; "Errore HTTP $ec. $d" }
        }
        $script:lblCreditRef.Text      = $msg
        $script:lblCreditRef.ForeColor = [System.Drawing.Color]::Red
        $script:txtCreditRef.Text      = ""
    } catch {
        $script:lblCreditRef.Text      = "Errore: $($_.Exception.Message)"
        $script:lblCreditRef.ForeColor = [System.Drawing.Color]::Red
        $script:txtCreditRef.Text      = ""
    }
})

$btnResetData.Add_Click({
    $confirm1 = [System.Windows.Forms.MessageBox]::Show(
        "Questa operazione cancellera' in modo permanente:`n`n" +
        "  - API Key Anthropic`n" +
        "  - Credenziali salvate (creds.xml)`n" +
        "  - Cache contatti GAL (contacts_cache.json)`n" +
        "  - Bozze locali (drafts.json)`n" +
        "  - Profilo di stile personale (style_*.json)`n`n" +
        "Continuare?",
        "Conferma reset dati personali",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirm1 -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $confirm2 = [System.Windows.Forms.MessageBox]::Show(
        "Sei sicuro? L'operazione non e' reversibile.",
        "Conferma definitiva",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirm2 -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $deleted = @()
    $errors  = @()

    # 1. Svuota API Key in memoria e nella UI
    $txtApiKey.Text        = ""
    $script:config.ApiKey  = ""
    $script:lastAIGenerated = $null

    # 2. Cancella file dati (incluso profilo di stile associato all'email corrente)
    $filesToDelete = @(
        @{ Path=$credsFile;          Name="Credenziali (creds.xml)" },
        @{ Path=$contactsFile;       Name="Cache contatti (contacts_cache.json)" },
        @{ Path=$draftsFile;         Name="Bozze locali (drafts.json)" },
        @{ Path=(Get-StyleFilePath); Name="Profilo di stile personale" }
    )

    foreach ($f in $filesToDelete) {
        if (Test-Path $f.Path) {
            try {
                Remove-Item $f.Path -Force
                $deleted += $f.Name
            } catch {
                $errors += "$($f.Name): $($_.Exception.Message)"
            }
        }
    }

    # 3. Aggiorna UI
    $script:contactsCache = @{}
    Refresh-DraftsList
    Refresh-StyleTab

    # 4. Report finale
    $msg = "Reset completato.`n`n"
    if ($deleted.Count -gt 0) { $msg += "Cancellati:`n" + ($deleted | ForEach-Object { "  - $_" } | Out-String) }
    if ($errors.Count -gt 0)  { $msg += "`nErrori:`n"   + ($errors  | ForEach-Object { "  - $_" } | Out-String) }
    $msg += "`nPuoi condividere lo script in sicurezza."

    [System.Windows.Forms.MessageBox]::Show($msg, "Reset completato",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
})
if ($script:config.ApiKey) {
    $txtApiKey.Text = $script:config.ApiKey
}

# ============================================================
# HELPER REPLY EDITOR - livello script (visibili dai closure)
# ============================================================
function Exec-REditCmd {
    param($cmd, $val="")
    if ($script:webRCompose -and $script:webRCompose.Document) {
        $script:webRCompose.Document.ExecCommand($cmd, $false, $val) | Out-Null
    }
}
function Get-RComposeHtml {
    try {
        $b = $script:webRCompose.Document.Body
        if ($b) { return $b.InnerHtml } else { return "" }
    } catch { return "" }
}
function Get-RComposePlain {
    try {
        $b = $script:webRCompose.Document.Body
        if ($b) { return $b.InnerText } else { return "" }
    } catch { return "" }
}

# ============================================================
# FUNZIONE OPEN-REPLYTAB (TAB DEDICATO PER RISPOSTE)
# ============================================================
function Open-ReplyTab {
    param($originalEmail, $replyAll = $false, $forward = $false)

    # Controlla se c'e' gia' un tab Reply/Forward aperto con testo non inviato
    $existing = $tab.TabPages | Where-Object { $_.Name -eq "tabReply" }
    if ($existing) {
        $hasContent = $false
        try {
            if ($script:webRCompose) {
                $existingText = Get-RComposePlain
                if (-not [string]::IsNullOrWhiteSpace($existingText)) { $hasContent = $true }
            }
        } catch {}

        if ($hasContent) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Hai una risposta/inoltro non ancora inviato con del testo scritto.`n`n" +
                "Aprendo una nuova risposta perderai il contenuto attuale non salvato.`n`n" +
                "Vuoi continuare e scartarlo?",
                "Bozza non salvata",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $tab.TabPages.Remove($existing)
    }

    # !! PRIMO: salva in script scope PRIMA di usarlo nei campi !!
    $script:rOrigEmail = $originalEmail

    # Crea nuovo tab reply
    $tabReply      = New-Object System.Windows.Forms.TabPage
    $tabReply.Name = "tabReply"
    $tabReply.Text = if ($forward) { "Inoltra" } elseif ($replyAll) { "Rispondi a tutti" } else { "Rispondi" }
    $tab.Controls.Add($tabReply)
    $tab.SelectedTab = $tabReply

    # TO -- in Reply: mittente originale; in ReplyAll: mittente + To originali; in Forward: vuoto
    $lblRTo = New-Label "To:" 10 15
    $tabReply.Controls.Add($lblRTo)
    $txtRTo = New-TextBox 80 12 870
    if ($forward) {
        $txtRTo.Text = ""
    } elseif ($replyAll) {
        # ReplyAll: mittente + tutti i To originali (uniti con ;)
        $allTo = @($originalEmail.From)
        if ($originalEmail.To) {
            $originalEmail.To -split ";" | ForEach-Object {
                $addr = $_.Trim()
                if ($addr -ne "" -and $addr -ne $originalEmail.From) { $allTo += $addr }
            }
        }
        $txtRTo.Text = ($allTo | Where-Object { $_ -ne "" }) -join "; "
    } else {
        # Reply semplice: solo il mittente
        $txtRTo.Text = $originalEmail.From
    }
    $tabReply.Controls.Add($txtRTo)

    # CC -- in ReplyAll: CC originali; altrimenti vuoto
    $lblRCc = New-Label "CC:" 10 45
    $tabReply.Controls.Add($lblRCc)
    $txtRCc = New-TextBox 80 42 870
    if ($replyAll -and $originalEmail.CC) {
        $txtRCc.Text = $originalEmail.CC
    }
    $tabReply.Controls.Add($txtRCc)

    # SUBJECT -- Re: o Fwd: + oggetto originale (evita doppio Re: Re:)
    $lblRSubj = New-Label "Oggetto:" 10 75
    $tabReply.Controls.Add($lblRSubj)
    $txtRSubj = New-TextBox 80 72 870
    $prefix   = if ($forward) { "Fwd: " } else { "Re: " }
    $subjOrig = if ($originalEmail.Subject) { $originalEmail.Subject } else { "" }
    $subjOrig = $subjOrig -replace "^\s*(Re|RE|Fwd|FWD|FW|Fw):\s*", ""  # rimuovi prefissi esistenti
    $txtRSubj.Text = "$prefix$subjOrig"
    $tabReply.Controls.Add($txtRSubj)
    
    # -- TOOLBAR REPLY (identica al Compose) ------------------
    $pnlRToolbar = New-Object System.Windows.Forms.Panel
    $pnlRToolbar.Location  = New-Object System.Drawing.Point(10, 105)
    $pnlRToolbar.Size      = New-Object System.Drawing.Size(940, 62)
    $pnlRToolbar.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 240)
    $tabReply.Controls.Add($pnlRToolbar)

    # Riga 1 - formattazione carattere
    $tbRB = New-TBtn "B"  4   3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold))
    $tbRI = New-TBtn "I"  35  3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Italic))
    $tbRU = New-TBtn "U"  66  3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Underline))
    $tbRS = New-TBtn "S"  97  3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Strikeout))

    $sepR1=New-Object System.Windows.Forms.Label; $sepR1.Location=New-Object System.Drawing.Point(130,4)
    $sepR1.Size=New-Object System.Drawing.Size(1,20); $sepR1.BorderStyle="FixedSingle"; $pnlRToolbar.Controls.Add($sepR1)

    $cmbRFont = New-Object System.Windows.Forms.ComboBox
    $cmbRFont.Location=New-Object System.Drawing.Point(135,4); $cmbRFont.Size=New-Object System.Drawing.Size(120,24)
    $cmbRFont.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbRFont.Font=New-Object System.Drawing.Font("Segoe UI",8)
    @("Aptos","Segoe UI","Arial","Calibri","Times New Roman","Courier New") |
        ForEach-Object { $cmbRFont.Items.Add($_) | Out-Null }
    $cmbRFont.SelectedIndex=0

    $cmbRSize = New-Object System.Windows.Forms.ComboBox
    $cmbRSize.Location=New-Object System.Drawing.Point(260,4); $cmbRSize.Size=New-Object System.Drawing.Size(50,24)
    $cmbRSize.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbRSize.Font=New-Object System.Drawing.Font("Segoe UI",8)
    @("8","9","10","11","12","14","16","18","20","24") |
        ForEach-Object { $cmbRSize.Items.Add($_) | Out-Null }
    $cmbRSize.SelectedIndex=4   # 12pt

    $sepR2=New-Object System.Windows.Forms.Label; $sepR2.Location=New-Object System.Drawing.Point(315,4)
    $sepR2.Size=New-Object System.Drawing.Size(1,20); $sepR2.BorderStyle="FixedSingle"; $pnlRToolbar.Controls.Add($sepR2)

    $btnRFgColor = New-TBtn "A" 319 3 28 26 (New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold))
    $btnRFgColor.ForeColor=[System.Drawing.Color]::Red
    $btnRBgColor = New-TBtn "ab" 350 3 32 26; $btnRBgColor.BackColor=[System.Drawing.Color]::Yellow

    # Bottoni azione riga 1 (destra)
    $btnRSend = New-Object System.Windows.Forms.Button
    $btnRSend.Text="INVIA"; $btnRSend.Location=New-Object System.Drawing.Point(530,3)
    $btnRSend.Size=New-Object System.Drawing.Size(100,26)
    $btnRSend.BackColor=[System.Drawing.Color]::SeaGreen; $btnRSend.ForeColor=[System.Drawing.Color]::White
    $btnRSend.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnRSend.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

    $btnRDraft = New-Object System.Windows.Forms.Button
    $btnRDraft.Text="Salva Bozza"; $btnRDraft.Location=New-Object System.Drawing.Point(635,3)
    $btnRDraft.Size=New-Object System.Drawing.Size(100,26)
    $btnRDraft.BackColor=[System.Drawing.Color]::SteelBlue; $btnRDraft.ForeColor=[System.Drawing.Color]::White
    $btnRDraft.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnRDraft.Font=New-Object System.Drawing.Font("Segoe UI",9)

    $btnRClose = New-Object System.Windows.Forms.Button
    $btnRClose.Text="Chiudi"; $btnRClose.Location=New-Object System.Drawing.Point(740,3)
    $btnRClose.Size=New-Object System.Drawing.Size(80,26)
    $btnRClose.BackColor=[System.Drawing.Color]::Tomato; $btnRClose.ForeColor=[System.Drawing.Color]::White
    $btnRClose.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnRClose.Font=New-Object System.Drawing.Font("Segoe UI",9)

    $btnRGenAI = New-Object System.Windows.Forms.Button
    $btnRGenAI.Text="Genera AI"; $btnRGenAI.Location=New-Object System.Drawing.Point(825,3)
    $btnRGenAI.Size=New-Object System.Drawing.Size(110,26)
    $btnRGenAI.BackColor=[System.Drawing.Color]::MediumPurple; $btnRGenAI.ForeColor=[System.Drawing.Color]::White
    $btnRGenAI.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnRGenAI.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

    # Riga 2 - allineamento + elenchi + rientro
    $tbRAlL  = New-TBtn "L"  4   33 28 26; $tbRAlL.ForeColor=[System.Drawing.Color]::DarkBlue
    $tbRAlC  = New-TBtn "C"  35  33 28 26; $tbRAlC.ForeColor=[System.Drawing.Color]::DarkBlue
    $tbRAlR  = New-TBtn "R"  66  33 28 26; $tbRAlR.ForeColor=[System.Drawing.Color]::DarkBlue
    $tbRAlJ  = New-TBtn "J"  97  33 28 26; $tbRAlJ.ForeColor=[System.Drawing.Color]::DarkBlue
    $tbRUl   = New-TBtn "ul" 135 33 30 26
    $tbROl   = New-TBtn "ol" 168 33 30 26
    $tbRIn   = New-TBtn ">>" 203 33 32 26
    $tbROut  = New-TBtn "<<" 238 33 32 26

    @($tbRB,$tbRI,$tbRU,$tbRS,$cmbRFont,$cmbRSize,$btnRFgColor,$btnRBgColor,
      $btnRSend,$btnRDraft,$btnRClose,$btnRGenAI,
      $tbRAlL,$tbRAlC,$tbRAlR,$tbRAlJ,$tbRUl,$tbROl,$tbRIn,$tbROut) |
        ForEach-Object { $pnlRToolbar.Controls.Add($_) }

    # -- WebBrowser editor risposta ----------------------------
    $webRCompose = New-Object System.Windows.Forms.WebBrowser
    $webRCompose.Location = New-Object System.Drawing.Point(10, 170)
    $webRCompose.Size     = New-Object System.Drawing.Size(940, 200)
    $webRCompose.IsWebBrowserContextMenuEnabled = $true
    $webRCompose.WebBrowserShortcutsEnabled     = $true
    $webRCompose.ScriptErrorsSuppressed         = $true
    $tabReply.Controls.Add($webRCompose)
    $script:webRCompose = $webRCompose

    $webRCompose.DocumentText = "<html><head><style>body{font-family:Aptos,'Segoe UI',sans-serif;font-size:12pt;line-height:1.15;margin:8px;color:#222;outline:none;}</style></head><body contenteditable='true' id='rbody'><br></body></html>"
    $webRCompose.Add_DocumentCompleted({
        try { $script:webRCompose.Document.Body.SetAttribute("contenteditable","true") } catch {}
    })

    # Salva in script scope per i closure degli handler
    $script:rTabRef    = $tabReply
    $script:rTo        = $txtRTo
    $script:rCc        = $txtRCc
    $script:rSubj      = $txtRSubj
    $script:rOrigEmail = $originalEmail
    $script:isForward  = $forward
    # Toolbar controls - tutti in script scope per PS5 closure
    $script:rCmbFont   = $cmbRFont
    $script:rCmbSize   = $cmbRSize
    $script:rFgBtn     = $btnRFgColor
    $script:rBgBtn     = $btnRBgColor

    # GAL autocomplete su To e CC
    Add-GalAutoComplete $txtRTo $tabReply
    Add-GalAutoComplete $txtRCc $tabReply

    # Handler toolbar formattazione reply - usa script: scope
    $tbRB.Add_Click({   Exec-REditCmd "bold" })
    $tbRI.Add_Click({   Exec-REditCmd "italic" })
    $tbRU.Add_Click({   Exec-REditCmd "underline" })
    $tbRS.Add_Click({   Exec-REditCmd "strikethrough" })
    $cmbRFont.Add_SelectedIndexChanged({
        if ($script:rCmbFont.SelectedItem) {
            Exec-REditCmd "fontName" $script:rCmbFont.SelectedItem.ToString()
        }
    })
    $cmbRSize.Add_SelectedIndexChanged({
        if ($script:rCmbSize.SelectedItem) {
            $pt = [int]$script:rCmbSize.SelectedItem.ToString()
            $fs = switch($pt){ {$_-le8}{1} {$_-le10}{2} {$_-le12}{3} {$_-le14}{4} {$_-le18}{5} {$_-le24}{6} default{7} }
            Exec-REditCmd "fontSize" $fs
        }
    })
    $btnRFgColor.Add_Click({
        $d = New-Object System.Windows.Forms.ColorDialog; $d.FullOpen=$true
        if ($d.ShowDialog() -eq "OK") {
            $h = "#{0:X2}{1:X2}{2:X2}" -f $d.Color.R,$d.Color.G,$d.Color.B
            $script:rFgBtn.ForeColor = $d.Color
            Exec-REditCmd "foreColor" $h
        }
    })
    $btnRBgColor.Add_Click({
        $d = New-Object System.Windows.Forms.ColorDialog; $d.FullOpen=$true
        if ($d.ShowDialog() -eq "OK") {
            $h = "#{0:X2}{1:X2}{2:X2}" -f $d.Color.R,$d.Color.G,$d.Color.B
            $script:rBgBtn.BackColor = $d.Color
            Exec-REditCmd "hiliteColor" $h
        }
    })
    $tbRAlL.Add_Click({ Exec-REditCmd "justifyLeft" })
    $tbRAlC.Add_Click({ Exec-REditCmd "justifyCenter" })
    $tbRAlR.Add_Click({ Exec-REditCmd "justifyRight" })
    $tbRAlJ.Add_Click({ Exec-REditCmd "justifyFull" })
    $tbRUl.Add_Click({  Exec-REditCmd "insertUnorderedList" })
    $tbROl.Add_Click({  Exec-REditCmd "insertOrderedList" })
    $tbRIn.Add_Click({  Exec-REditCmd "indent" })
    $tbROut.Add_Click({ Exec-REditCmd "outdent" })

    # -- SPLIT: email originale sotto -------------------------
    $lblOrig = New-Label "--- Messaggio originale da: $($script:rOrigEmail.FromName) - $($script:rOrigEmail.Date) ---" 10 380 920 20
    $lblOrig.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblOrig.ForeColor = [System.Drawing.Color]::Gray
    $tabReply.Controls.Add($lblOrig)

    $webOrig = New-Object System.Windows.Forms.WebBrowser
    $webOrig.Location = New-Object System.Drawing.Point(10, 400)
    $webOrig.Size     = New-Object System.Drawing.Size(940, 230)
    $webOrig.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor
                        [System.Windows.Forms.AnchorStyles]::Left -bor
                        [System.Windows.Forms.AnchorStyles]::Right -bor
                        [System.Windows.Forms.AnchorStyles]::Bottom
    $webOrig.ScrollBarsEnabled = $true
    $webOrig.IsWebBrowserContextMenuEnabled = $false
    $webOrig.AllowWebBrowserDrop = $false
    $webOrig.ScriptErrorsSuppressed = $true
    $tabReply.Controls.Add($webOrig)

    # Carica HTML originale
    $tempReplyPath = "$rootPath\reply_temp.html"
    $oe = $script:rOrigEmail
    if ($oe.HTMLBody -and $oe.HTMLBody.Trim() -ne "") {
        [System.IO.File]::WriteAllText($tempReplyPath, $oe.HTMLBody, [System.Text.Encoding]::UTF8)
        $webOrig.Navigate("file:///$($tempReplyPath.Replace('\','/'))")
    } else {
        $plain = $oe.Body -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace "`n","<br>"
        $html  = "<html><head><style>body{font-family:'Segoe UI';font-size:9pt;color:#666;margin:10px;}</style></head><body>$plain</body></html>"
        [System.IO.File]::WriteAllText($tempReplyPath, $html, [System.Text.Encoding]::UTF8)
        $webOrig.Navigate("file:///$($tempReplyPath.Replace('\','/'))")
    }
    
    $btnRGenAI.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:config.ApiKey)) {
            [System.Windows.Forms.MessageBox]::Show("Configura la API Key nel tab Impostazioni.", "API Key mancante",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $isFwd = $script:isForward
        $dlgTitle = if ($isFwd) { "Componi testo di inoltro con AI" } else { "Genera risposta con AI" }
        $defPrompt = if ($isFwd) { "Scrivi un breve testo introduttivo per presentare il messaggio inoltrato" } else { "" }

        $opts = Show-AIDialog -title $dlgTitle -defaultPrompt $defPrompt
        if ($null -eq $opts) { return }

        $oe = $script:rOrigEmail
        $emailCtx = "Da: $($oe.FromName) <$($oe.From)>`nOggetto: $($oe.Subject)`nData: $($oe.Date)`n`n$($oe.Body)"

        $p = if ([string]::IsNullOrWhiteSpace($opts.Prompt)) {
            if ($isFwd) { "Scrivi un breve testo introduttivo per presentare il messaggio inoltrato" }
            else        { "Rispondi in modo appropriato" }
        } else { $opts.Prompt }

        try {
            if ($isFwd) {
                # FORWARD: il prompt e' istruzione principale, il corpo e' contesto/riferimento
                $result = Invoke-AI -prompt $p -apiKey $script:config.ApiKey `
                                    -replyContext $emailCtx `
                                    -language $opts.Language -tone $opts.Tone `
                                    -mode "forward"
            } else {
                # REPLY: genera risposta all'email originale
                $result = Invoke-AI -prompt $p -apiKey $script:config.ApiKey `
                                    -replyContext $emailCtx `
                                    -language $opts.Language -tone $opts.Tone `
                                    -mode "reply"
            }
            try {
            $b = $script:webRCompose.Document.Body
            if ($b) {
                $html = ($result -replace "`r`n","<br>" -replace "`n","<br>")
                $b.InnerHtml = $html
                $b.SetAttribute("contenteditable","true")
            }
        } catch { }
            $script:lastAIGenerated = $result.Trim()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Errore AI:`n$_", "Errore",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Handler: Invia
    $btnRSend.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:rTo.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Inserisci almeno un destinatario.", "Attenzione",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $finalPlainR = Get-RComposePlain
        if ([string]::IsNullOrWhiteSpace($finalPlainR)) {
            [System.Windows.Forms.MessageBox]::Show("Il corpo del messaggio e' vuoto.", "Attenzione",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Apprendimento stile
        if ($script:appSettings.StyleLearningEnabled -and $script:lastAIGenerated -and -not [string]::IsNullOrWhiteSpace($script:lastAIGenerated)) {
            try { Update-StyleProfile -aiText $script:lastAIGenerated -userText $finalPlainR -apiKey $script:config.ApiKey } catch {}
            $script:lastAIGenerated = $null
        }

        $oe = $script:rOrigEmail
        # Corpo risposta plain text -> HTML
        $replyHtml = Get-RComposeHtml
        # Email originale come blocco HTML sotto la risposta
        $origHtml  = if ($oe.HTMLBody) { $oe.HTMLBody } else {
            ($oe.Body -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace "`n","<br>")
        }
        $fullBody = "<html><body style='font-family:Aptos,Segoe UI,sans-serif;font-size:12pt;'>" +
                    "$replyHtml" +
                    "<br><br><div style='border-top:1px solid #ccc;padding-top:6px;color:#555;font-size:9pt;'>" +
                    "<b>Da:</b> $($oe.FromName) &lt;$($oe.From)&gt;<br>" +
                    "<b>Inviato:</b> $($oe.Date)<br>" +
                    "<b>A:</b> $($oe.To)<br>" +
                    $(if ($oe.CC) { "<b>Cc:</b> $($oe.CC)<br>" } else { "" }) +
                    "<b>Oggetto:</b> $($oe.Subject)<br><br>" +
                    "$origHtml</div></body></html>"
        if (Send-OutlookEmail -to $script:rTo.Text -cc $script:rCc.Text -subject $script:rSubj.Text -body $fullBody -attachments @()) {
            $tab.TabPages.Remove($script:rTabRef)
        }
    })
    
    # Handler: Salva Bozza
    $btnRDraft.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:rSubj.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Inserisci un oggetto.", "Attenzione",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        Save-Draft -to $script:rTo.Text -cc $script:rCc.Text -subject $script:rSubj.Text -body (Get-RComposeHtml)
        Refresh-DraftsList
        [System.Windows.Forms.MessageBox]::Show("Bozza salvata nel tab Bozze!", "Successo",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    
    # Handler: Chiudi
    $btnRClose.Add_Click({
        $tab.TabPages.Remove($script:rTabRef)
    })
}

# ============================================================
# SHOW DIALOG
# ============================================================

# ============================================================
# FUNZIONE REFRESH INBOX (riutilizzata da splash e da btnRefresh)
# ============================================================
function Invoke-InboxRefresh {
    param($progressCallback = $null)

    $maxEm = if ($script:appSettings.ContainsKey('MaxEmails')) { [int]$script:appSettings.MaxEmails } else { 30 }

    $onProg = {
        param($current, $total)
        if ($progressCallback) {
            try { & $progressCallback $current $total } catch {}
        }
    }

    $script:inboxEmails = Get-OutlookEmails -maxMessages $maxEm -onProgress $onProg

    # Aggiorna cache contatti IN MEMORIA (una sola lettura/scrittura disco, non una per contatto)
    $contactMap = Load-Contacts
    function Add-ContactInMemory {
        param($map, $email, $name)
        if (-not $email -or $email -notmatch "@") { return }
        $email = $email.Trim().ToLower()
        $name  = if ($name -and $name -ne $email) { $name.Trim() } else { $email }
        if ($map[$email]) {
            $map[$email].count = [int]$map[$email].count + 1
            $map[$email].last  = (Get-Date).ToString("yyyy-MM-dd")
            if ($name -and $name -ne $email) { $map[$email].name = $name }
        } else {
            $map[$email] = [PSCustomObject]@{ name=$name; count=1; last=(Get-Date).ToString("yyyy-MM-dd") }
        }
    }

    foreach ($e in $script:inboxEmails) {
        try {
            if ($e.From -and $e.From -match "@") { Add-ContactInMemory $contactMap $e.From $e.FromName }
            if ($e.To) {
                $e.To -split ";" | ForEach-Object {
                    $addr = $_.Trim()
                    if ($addr -match "<(.+@.+)>") { Add-ContactInMemory $contactMap $Matches[1] ($addr -replace "<.+>","").Trim() }
                    elseif ($addr -match "@") { Add-ContactInMemory $contactMap $addr $addr }
                }
            }
            if ($e.CC) {
                $e.CC -split ";" | ForEach-Object {
                    $addr = $_.Trim()
                    if ($addr -match "<(.+@.+)>") { Add-ContactInMemory $contactMap $Matches[1] ($addr -replace "<.+>","").Trim() }
                    elseif ($addr -match "@") { Add-ContactInMemory $contactMap $addr $addr }
                }
            }
        } catch {}
    }
    Save-Contacts $contactMap
    $script:contactsCache = $contactMap
    $lstEmails.Items.Clear()
    foreach ($e in $script:inboxEmails) {
        $stato = if ($e.Unread) { "Non letto" } else { "Letto" }
        $item  = New-Object System.Windows.Forms.ListViewItem($e.FromName)
        $item.SubItems.Add($e.Subject) | Out-Null
        $item.SubItems.Add($e.Date)    | Out-Null
        $item.SubItems.Add($stato)     | Out-Null
        if ($e.Unread) {
            $item.ForeColor = [System.Drawing.Color]::FromArgb(0,100,200)
            $item.Font      = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
        }
        $lstEmails.Items.Add($item) | Out-Null
    }
    $lblStatus.Text = "$($script:inboxEmails.Count) email caricate"
    $txtReadSubj.Text = ""; $txtReadFrom.Text = ""
    $script:pendingHTML = ""
    $webBody.DocumentText = "<html><body></body></html>"
    $lstAttachRead.Items.Clear()
    $btnReply.Enabled = $false; $btnReplyAll.Enabled = $false; $btnForward.Enabled = $false
}

$form.Add_Shown({
    $form.Activate()
    $picLogo.Left = $pnlHeader.Width - $picLogo.Width - 3
    Refresh-DraftsList
    Refresh-StyleTab

    # Applica impostazioni aspetto salvate
    $pnlHeader.BackColor = Parse-Color $script:appSettings.HeaderColor
    $form.BackColor      = Parse-Color $script:appSettings.FormBgColor
    $lblAppTitle.Text    = $script:appSettings.AppTitle
    $form.Text           = $script:appSettings.FormTitle
    # Font e colore MOTD
    $motdFont  = if ($script:appSettings.MotdFont) { $script:appSettings.MotdFont } else { "Segoe UI" }
    $motdSize  = if ($script:appSettings.MotdSize) { [float]$script:appSettings.MotdSize } else { 14 }
    $motdColor = if ($script:appSettings.MotdColor) { Parse-Color $script:appSettings.MotdColor } else { [System.Drawing.Color]::FromArgb(30,30,30) }
    $lblAppTitle.Font      = New-Object System.Drawing.Font($motdFont, $motdSize, [System.Drawing.FontStyle]::Bold)
    $lblAppTitle.ForeColor = $motdColor

    # Vai al tab Posta in Arrivo
    $tab.SelectedTab = $tabInbox

    # -- SPLASH SCREEN -----------------------------------------
    $splash = New-Object System.Windows.Forms.Form
    $splash.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $splash.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $splash.Size            = New-Object System.Drawing.Size(420, 340)
    $splash.BackColor       = [System.Drawing.Color]::White
    $splash.TopMost         = $true
    $splash.ShowInTaskbar   = $false

    # Bordo giallo Fastweb
    $splash.Add_Paint({
        param($s,$e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255,216,0), 4)
        $e.Graphics.DrawRectangle($pen, 2, 2, $s.Width-5, $s.Height-5)
        $pen.Dispose()
    })

    # Logo centrato
    $picS = New-Object System.Windows.Forms.PictureBox
    $picS.Size     = New-Object System.Drawing.Size(180, 180)
    $picS.Location = New-Object System.Drawing.Point(120, 30)
    $picS.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $picS.BackColor = [System.Drawing.Color]::White
    $picS.Image = Load-Logo -w 180 -h 180
    $splash.Controls.Add($picS)

    # Label stato
    $lblSplash = New-Object System.Windows.Forms.Label
    $lblSplash.Text      = "Caricamento posta in arrivo..."
    $lblSplash.Location  = New-Object System.Drawing.Point(10, 220)
    $lblSplash.Size      = New-Object System.Drawing.Size(400, 22)
    $lblSplash.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblSplash.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblSplash.ForeColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $splash.Controls.Add($lblSplash)

    # Barra di progresso
    $prog = New-Object System.Windows.Forms.ProgressBar
    $prog.Location = New-Object System.Drawing.Point(20, 252)
    $prog.Size     = New-Object System.Drawing.Size(378, 22)
    $prog.Minimum  = 0
    $prog.Maximum  = 100
    $prog.Value    = 0
    $prog.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $prog.ForeColor = [System.Drawing.Color]::FromArgb(255,216,0)
    $splash.Controls.Add($prog)

    # Copyright / versione
    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text      = "Copezzot Mail AI Assistant  v20"
    $lblVer.Location  = New-Object System.Drawing.Point(10, 290)
    $lblVer.Size      = New-Object System.Drawing.Size(400, 18)
    $lblVer.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblVer.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblVer.ForeColor = [System.Drawing.Color]::Gray
    $splash.Controls.Add($lblVer)

    $splash.Show($form)
    $splash.Refresh()

    # Callback progresso reale: aggiornato per ogni email effettivamente scaricata
    $splashProgress = {
        param($current, $total)
        if ($total -gt 0) {
            $pct = [int](($current / $total) * 100)
            $prog.Value     = [Math]::Min($pct, 100)
            $lblSplash.Text = "Caricamento email $current di $total..."
            $splash.Refresh()
        }
    }

    # Esegui refresh inbox con progresso reale
    try {
        Invoke-InboxRefresh -progressCallback $splashProgress
    } catch { }

    # Completa la barra e chiudi splash
    $prog.Value      = 100
    $lblSplash.Text  = "Pronto  -  $($script:inboxEmails.Count) email caricate"
    $splash.Refresh()
    Start-Sleep -Milliseconds 250
    $splash.Close()
    $splash.Dispose()
})

[void]$form.ShowDialog()
