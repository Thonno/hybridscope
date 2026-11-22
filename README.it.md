<div align="center">

# HybridScope

**Una sola vista sulle identità ibride: Active Directory + Exchange on-premises + Microsoft 365.**

Chi è on-prem e chi è già nel cloud, quanto spazio occupa ogni mailbox, chi ha licenza, chi ha l'MFA, chi non accede da mesi: tutto in una finestra, esportabile in Excel.

[![Sponsor su GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Offrimi un caffè](https://img.shields.io/badge/Offrimi%20un%20caff%C3%A8-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/aascione956)
[![Metti una stella](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Stella)](https://github.com/Thonno/hybridscope/stargazers)

![PowerShell 5.1](https://img.shields.io/badge/PowerShell-5.1-5391FE?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server-0078D6?logo=windows&logoColor=white)
![Licenza: MIT](https://img.shields.io/badge/Licenza-MIT-green.svg)
![Lingue](https://img.shields.io/badge/UI-English%20%7C%20Italiano-blue)

[English](README.md) · **Italiano**

<img src="docs/07-overview.png" alt="Panoramica HybridScope" width="900">

</div>

---

## Indice

- [Funzionalità](#funzionalità)
- [Screenshot](#screenshot)
- [Requisiti](#requisiti)
- [Avvio rapido](#avvio-rapido)
- [Uso dell'interfaccia](#uso-dellinterfaccia)
- [Tab e colonne](#tab-e-colonne)
- [Configurazione (config.json)](#configurazione-configjson)
- [Permessi](#permessi)
- [Riga di comando (senza interfaccia)](#riga-di-comando-senza-interfaccia)
- [Report pianificati (-Auto)](#report-pianificati--auto)
- [Risoluzione problemi](#risoluzione-problemi)
- [FAQ](#faq)
- [Roadmap](#roadmap)
- [Sostieni il progetto](#sostieni-il-progetto)
- [Ringraziamenti](#ringraziamenti)
- [Licenza](#licenza)

## Funzionalità

- **Interfaccia stile RVTools**: ogni insieme di dati nel suo tab, ricerca che filtra il tab corrente, colonne ordinabili, copia con intestazioni.
- **Tre fonti in un unico report**:
  - **Active Directory**: utenti, OU, stato, ultimo logon, ultimo cambio password
  - **Exchange Server on-premises**: mailbox, dimensione, elementi, database, quota, ultimo accesso, e le mailbox remote già migrate
  - **Microsoft 365 (Microsoft Graph)**: licenze, ultimo accesso, mailbox Exchange Online, OneDrive, registrazione MFA e metodi, stato MFA per-user, licenze acquistate e libere
- **Posizione della mailbox affidabile**: `ON-PREM` / `365` / `NO MAILBOX` arriva da Exchange (`Get-Mailbox`, `Get-RemoteMailbox`) e da Graph (`provisionedPlans`), mai dallo spazio occupato.
- **Stato MFA reale**: se è registrato almeno un metodo, quali metodi, quello predefinito e lo stato per-user legacy.
- **Scheda utente**: doppio clic su un utente in qualsiasi tab per vedere in una sola finestra AD, mailbox, OneDrive, MFA, licenze e dati di sincronizzazione, insieme ai suoi **gruppi e ruoli**: gruppi AD (diretti, annidati e primario), gruppi Entra (Microsoft 365, sicurezza, distribuzione, dinamici) e ruoli Entra.
- **Controllo della sincronizzazione (tab SYNC)**: confronta AD ed Entra ID come li collegano Entra Connect / Cloud Sync e segnala cosa è rotto o sta per rompersi: `ImmutableId` che non corrisponde più all'ancora AD, account sincronizzati che non esistono più in AD, account cloud con un `ImmutableId` residuo, account AD che alla prima sincronizzazione faranno soft-match o andranno in conflitto con un account cloud, errori di provisioning, UPN finiti su `onmicrosoft.com`, e quando è girata l'ultima sincronizzazione.
- **Panoramica** per gruppo (OU) con totali, più una lista **Da attenzionare**:
  - mailbox on-prem da migrare, ordinate per dimensione
  - utenti licenziati senza MFA
  - licenze assegnate ad account disabilitati
  - account inattivi
  - account AD non sincronizzati
  - problemi di sincronizzazione
- **Export**: XLSX formattato o CSV. **FULL** (tutti i tab), un **preset** (accessi, licenze, migrazione, MFA, anagrafica) oppure **personalizzato**: spunti solo le colonne che vuoi. Puoi salvare i tuoi preset. Oppure solo il tab corrente con il filtro applicato.
- **Apri un report salvato** (File > Apri report, Ctrl+O, oppure trascinalo sulla finestra): un report XLSX, una cartella di export CSV o un singolo CSV, in entrambe le lingue. Sfogli, cerchi, riesporti e apri le schede utente senza collegarti a niente, per esempio l'`HybridScope_LATEST.xlsx` giornaliero da un altro PC.
- **Funziona anche da PC fuori dominio**: credenziali Windows oppure utente e password di dominio inseriti a mano, con Kerberos.
- **Modalità di infrastruttura**: *Full Cloud*, *Full Hybrid*, *Hybrid*, *On-Prem* o *AD*, selezionabili con un clic. L'app mostra e unisce solo le fonti che hai.
- **Raggruppamento in tab** scelto dalla finestra: OU di primo livello, OU che contiene l'utente, lista unica oppure regole personalizzate.
- **Indicatori di avanzamento**: Connetti 365 › Connetti Exchange › Verifica AD › Carica dati diventano verdi con una ✓ quando sono completati, gialli durante l'esecuzione e rossi in caso di errore.
- Pulsante **Cambia account** per disconnettersi da Microsoft 365 e cancellare l'accesso salvato, e pulsante **Interrompi** per fermare un caricamento in corso.
- **Riga di comando**: `-Auto` per il report completo, `-Only MFA` (oppure `Licenses`, `Mailbox365`, `OneDrive`, `MailboxOnPrem`, `Sync`) per un solo insieme di dati, `-PassThru` per avere oggetti nella pipeline. Accesso con certificato per i report pianificati. Mantiene una copia `HybridScope_LATEST.xlsx` e uno storico dell'andamento (utenti on-prem, spazio, % MFA).
- Menu **File** e **Aiuto**: apri report, export, cartella dei report, documentazione (F1), **Cerca aggiornamenti** (confronta con l'ultima release su GitHub), segnalazione problemi, finestra **Informazioni**.
- Interfaccia in **italiano / inglese**.
- **Sola lettura**: HybridScope non modifica nulla in AD, Exchange o Microsoft 365.

## Screenshot

**Panoramica**: conteggi per OU / gruppo, posizione delle mailbox, spazio, licenze, MFA e inattività a colpo d'occhio.

![Panoramica](docs/07-overview.png)

**Utenti**: una riga per utente, con i dati di AD, Exchange e Microsoft 365 uniti. Cerca, ordina, esporta.

![Utenti](docs/08-users.png)

**Modalità di infrastruttura**: un clic e la finestra mostra solo le connessioni che servono.

| Full Hybrid | Full Cloud |
|---|---|
| ![Full Hybrid](docs/01-full-hybrid.png) | ![Full Cloud](docs/02-full-cloud.png) |
| **Hybrid** | **On-Prem** |
| ![Hybrid](docs/03-hybrid.png) | ![On-Prem](docs/04-on-prem.png) |
| **AD** | **Connesso, caricamento in corso** |
| ![AD](docs/05-ad.png) | ![Caricamento](docs/06-connected-loading.png) |

<sub>I dati sensibili (utenti, domini, server, nomi delle OU) sono sfocati.</sub>

## Requisiti

| | |
|---|---|
| Sistema | Windows 10 / 11 o Windows Server 2016+ |
| PowerShell | **Windows PowerShell 5.1**, già incluso in Windows |
| Moduli | `Microsoft.Graph.Authentication`, `ImportExcel`, `ActiveDirectory` (RSAT). All'avvio HybridScope propone di installare quelli mancanti. |
| Rete | HTTPS verso Microsoft 365; HTTP 80 (WinRM) verso il server Exchange; 9389 (ADWS) verso un domain controller |
| Exchange | Exchange Server 2013 / 2016 / 2019 / SE. Facoltativo: senza, ottieni solo AD + Microsoft 365. |
| Microsoft 365 | Ultimo accesso e report di registrazione MFA richiedono **Entra ID P1** (incluso in Business Premium, E3, E5). Senza P1, HybridScope passa a un metodo più lento e lo segnala. |

## Avvio rapido

1. **Scarica** l'ultima release (o fai `git clone`) in una cartella locale, es. `C:\HybridScope`. Evita OneDrive e Desktop.
2. **Abilita gli script locali** una volta sola. Da PowerShell:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   Get-ChildItem C:\HybridScope -Recurse | Unblock-File
   ```
3. **Avvia** con doppio clic su **`Start-HybridScope.cmd`**, oppure da PowerShell:
   ```powershell
   cd C:\HybridScope
   .\HybridScope.ps1
   ```
4. Se te lo propone, lascia che installi i moduli mancanti. Per `ActiveDirectory` si apre una finestra PowerShell visibile che chiede la conferma di amministratore.

### Facoltativo: creare un .exe

```powershell
.\tools\Build-Exe.ps1
```
Crea `HybridScope.exe` nella cartella principale con [PS2EXE](https://github.com/MScholtes/PS2EXE) e installa PS2EXE se manca. L'exe deve restare nella stessa cartella di `tools\` e `config.json`.

> ⚠️ Gli antivirus e le sandbox (Defender, Bitdefender...) segnalano spesso gli exe creati con PS2EXE come falsi positivi. Firma l'exe con un certificato di code signing (il comando viene mostrato alla fine della build), oppure usa semplicemente `Start-HybridScope.cmd`.

Al primo avvio viene creato `config.json` accanto allo script. `config.sample.json` mostra tutte le opzioni.

## Uso dell'interfaccia

0. **Infrastruttura**: in alto nella finestra clicca il pulsante che corrisponde al tuo ambiente: *Full Cloud*, *Full Hybrid*, *Hybrid*, *On-Prem* o *AD*. Vengono mostrati solo i riquadri che servono.
1. **Microsoft 365**: inserisci il tenant se vuoi (`contoso.onmicrosoft.com`), poi **Connetti 365**. Si apre il browser. La prima volta un amministratore deve dare il consenso ai permessi, tutti in sola lettura.
2. **Credenziali di dominio**:
   - lascia **Usa credenziali Windows** su un PC nel dominio;
   - altrimenti togli la spunta e scrivi utente e password: `mario@contoso.local`, `CONTOSO\mario` oppure solo `mario`, e in quel caso `@dominio` viene aggiunto in automatico.

   Valgono sia per Exchange sia per AD. L'utente viene ricordato, la password mai.

   Sei entrato in Microsoft 365 con l'account sbagliato? Premi **Cambia account**: disconnette e cancella l'accesso salvato, poi **Connetti 365** chiede di nuovo l'account. L'accesso non viene mai conservato tra una sessione e l'altra.
3. **Exchange on-premises**: inserisci il nome **interno** di un server Exchange (es. `exch01.contoso.local`) e scegli la modalità:
   - `http` + `Kerberos`: consigliata, nessun certificato
   - `https`: solo con un certificato valido per quel nome
   - `local`: se HybridScope gira direttamente sul server Exchange

   Poi premi **Connetti Exchange**.
4. **Active Directory**: controlla il dominio e premi **Verifica AD**.
5. Scegli **Raggruppa per**:
   - *OU di primo livello*: la prima OU sotto il dominio, es. `Vendite`
   - *OU di appartenenza*: l'OU che contiene direttamente l'utente, es. `Interni` o `Esterni`; di solito è la più utile
   - *Nessun raggruppamento*: una lista unica
   - *Regole personalizzate*: `GroupRules` in `config.json`

   Non devi mai mappare le OU a mano: il raggruppamento è automatico e la scelta viene ricordata.
   In **Contenuto export** scegli cosa conterrà l'export: **FULL**, un preset oppure solo le colonne che spunti (**Colonne...**). Vedi [Export](#export-full-preset-o-colonne-a-scelta).
6. Premi **Carica dati**. L'avanzamento compare nel pannello **Log**; qualche centinaio di utenti richiede pochi minuti. **Interrompi** ferma il caricamento; attende la fine della richiesta in corso.
7. Cerca, ordina ed **esporta**. Per guardare un report salvato in precedenza usa **File > Apri report** (Ctrl+O) o trascina il file `.xlsx` / la cartella CSV sulla finestra. **Doppio clic su un utente** in qualsiasi tab (UTENTI, un gruppo, MFA, mailbox, DA_ATTENZIONARE, SYNC...) per aprire la sua scheda.

La lingua si cambia dal menu in alto a destra (serve un riavvio). **Aiuto > Cerca aggiornamenti** ti dice se è uscita una versione più recente; **Aiuto > Informazioni su HybridScope** mostra la versione e i link del progetto.

### Modalità di infrastruttura

Scegli la tua infrastruttura con i pulsanti in alto nella finestra. Riquadri di connessione, tab, colonne e panoramica si adattano alle fonti che hai. Per cambiare basta un clic: la scelta non è legata ai dati.

| Pulsante | Fonti | Cosa ottieni |
|---|---|---|
| **Full Cloud** | Microsoft 365 | Utenti, licenze, mailbox Exchange Online, OneDrive, MFA, ultimo accesso |
| **Full Hybrid** | Microsoft 365 + Exchange on-prem + AD | Tutto: posizione on-prem o 365, mailbox su entrambi i lati, MFA, licenze, dati AD |
| **Hybrid** | Microsoft 365 + AD | Utenti AD uniti ai loro account cloud: licenze, Exchange Online, OneDrive, MFA |
| **On-Prem** | Exchange on-prem + AD | Utenti AD con le loro mailbox on-prem: dimensione, database, quota, ultimo accesso, inattività |
| **AD** | Active Directory | Utenti, OU, stato, ultimo logon, età della password, inattività |

Per i report pianificati (`-Auto`) la modalità si imposta con `Mode` in `config.json`.

### Export: FULL, preset o colonne a scelta

Scegli **cosa esportare prima di caricare i dati**, nel riquadro **Contenuto export** accanto alle connessioni (la scelta viene salvata e si può cambiare anche dopo il caricamento). **Esporta XLSX...** / **Esporta CSV...** scrivono poi esattamente quello, senza chiedere altro.

Nel riquadro, il menu a tendina sceglie il contenuto e **Colonne...** apre la scelta delle colonne (compaiono solo le colonne della modalità di infrastruttura selezionata):

- **FULL**: tutti i tab con tutte le colonne, come prima.
- **Un preset**, pronto per le richieste più comuni:

  | Preset | Colonne della lista utenti |
  |---|---|
  | Accessi e inattività | nome, cognome, username, stato, ultimo logon AD, ultimo accesso Microsoft 365, ultimo cambio password, licenze |
  | Licenze | nome, cognome, username, UPN, stato, licenze, ultimo accesso Microsoft 365, note |
  | Migrazione mailbox | gruppo, nome, cognome, username, stato, Location, tipo mailbox, dimensione on-prem / 365 / OneDrive |
  | Sicurezza MFA | nome, cognome, username, UPN, stato, licenze, stato MFA, metodo predefinito, metodi, stato per-user |
  | Anagrafica | gruppo, nome, cognome, username, UPN, OU, stato, origine |

- **Personalizzato**: spunti solo le colonne che ti servono, per esempio solo *AD_Username*, *AD_UltimoLogon* e *M365_UltimoAccesso*. Se cambi una colonna di un preset si passa a Personalizzato; **Salva come preset...** conserva la scelta con un nome a piacere (in `config.json`, mostrato con una ★).

Opzioni per preset e personalizzato: **un foglio per gruppo**, **includi PANORAMICA**, **solo account abilitati**. Un CSV con la sola lista utenti è un unico file; con più fogli ottieni una cartella con un CSV per tab. Accanto al menu vedi cosa verrà esportato, es. *3 colonne: AD_Username, AD_UltimoLogon, M365_UltimoAccesso*.

## Tab e colonne

| Tab | Contenuto |
|---|---|
| **PANORAMICA** | Conteggi per gruppo e **TOTALE**: utenti, posizione mailbox, spazio (GB), licenze, MFA, inattività |
| **UTENTI** | Tutti gli utenti, una riga ciascuno |
| *un tab per gruppo* | Utenti di ogni OU / gruppo (vedi `GroupBy` e `GroupRules`) |
| **ALTRI** | Utenti AD che non corrispondono a nessuna regola |
| **SOLO_CLOUD** | Utenti Entra senza account AD |
| **MAILBOX_ONPREM** | Mailbox on-prem: tipo, database, dimensione, elementi, eliminati, ultimo accesso, quota |
| **MAILBOX_365** | Mailbox Exchange Online dal report di utilizzo Microsoft |
| **ONEDRIVE** | Account OneDrive: spazio usato, quota, file, ultima attività |
| **MFA** | Ogni membro Entra: registrato, utilizzabile, metodo predefinito, metodi, stato per-user, SSPR, admin |
| **LICENZE_TENANT** | Licenze acquistate / assegnate / libere / sospese per SKU |
| **SYNC** | Controllo della sincronizzazione AD ↔ Entra ID (modalità con AD e Microsoft 365); vedi [Controllo della sincronizzazione](#controllo-della-sincronizzazione) |
| **DA_ATTENZIONARE** | Lista di azioni: mailbox da migrare, licenziati senza MFA, licenziati disabilitati, account inattivi, account AD non sincronizzati, problemi di sincronizzazione |
| **STORICO** | Una riga per ogni esecuzione `-Auto`, per seguire l'andamento |

<details>
<summary><b>Colonne utente</b> (UTENTI e tab dei gruppi)</summary>

| Colonna | Descrizione |
|---|---|
| Gruppo | Gruppo / tab OU dell'utente |
| Nome, Cognome, AD_Username | Da AD (o da Entra per gli utenti solo cloud) |
| OU | Percorso dell'OU dell'utente, es. `CONTOSO\Vendite\Dipendenti` |
| Stato | Abilitato / Disabilitato |
| AD_UltimoLogon | `lastLogonTimestamp`, replicato, preciso a circa 14 giorni |
| M365_UltimoAccesso | Ultimo accesso interattivo o non interattivo |
| Licenze | Nomi leggibili delle licenze (vedi `LicenseNames`) |
| Mailbox365_MB / MailboxOnPrem_MB / OneDrive_MB | Spazio usato in MB (2 decimali). Vuoto se Microsoft non riporta la dimensione |
| UltimoCambioPassword | Ultimo cambio password (AD) |
| **Location** | `365`, `ON-PREM` o `NO MAILBOX`; vedi sotto |
| TipoMailbox | `UserMailbox`, `SharedMailbox`, `RemoteUserMailbox`, `Exchange Online`... |
| FonteMailbox | Da dove arriva l'informazione sulla posizione |
| MFA_Stato | Configurato / Non configurato / Non presente in Entra |
| MFA_MetodoDefault, MFA_Metodi | Metodo MFA predefinito e tutti i metodi registrati |
| MFA_PerUser | MFA per-user legacy: `disabled` / `enabled` / `enforced` |
| UPN | User principal name di Microsoft 365 |
| Origine | AD sincronizzato / Solo AD / AD + cloud (non collegati) / Solo cloud / Guest. *Non collegati*: un account AD e un account cloud con lo stesso nome, ma Entra non sincronizza l'uno dall'altro |
| ProblemiSync | Problemi di sincronizzazione dell'utente (dettagli nel tab SYNC) |
| Note | Anomalie, es. "Licenziato senza MFA" o "Password senza scadenza" |
</details>

**Come viene decisa la Location:**
1. Esiste una mailbox su Exchange on-prem (`Get-Mailbox`): **ON-PREM**
2. Esiste una `RemoteMailbox` su Exchange on-prem, cioè la mailbox è stata migrata: **365**
3. Exchange Online è attivo per l'utente (Graph `provisionedPlans`) oppure l'utente compare nel report di utilizzo: **365**
4. Nessuno dei casi precedenti: **NO MAILBOX**

### Controllo della sincronizzazione

Il tab SYNC confronta AD ed Entra ID come fanno Entra Connect e Cloud Sync: tramite l'**ancora di origine** (source anchor). In AD l'ancora è `ms-DS-ConsistencyGuid` (oppure `objectGUID` se è vuoto); in Entra è salvata, in base64, in `onPremisesImmutableId`. Anche negli altri tab gli utenti vengono abbinati prima per ancora e poi, in mancanza, per nome.

| Categoria | Cosa significa |
|---|---|
| 0 - Stato della sincronizzazione | Quando è girata l'ultima sincronizzazione; segnalata se è più vecchia di `SyncStaleHours`, o se la sincronizzazione non è attiva nel tenant |
| 1 - ImmutableId diverso dall'ancora AD | Stesso nome account, ma Entra è collegato a un altro oggetto. Di solito l'account AD è stato cancellato e ricreato |
| 2 - Account sincronizzato non trovato in AD | Sincronizzato dal dominio / OU che leggi, ma nessun account AD ha la sua ancora |
| 3 - Account cloud con ImmutableId residuo | Sincronizzazione disattivata per l'account ma `ImmutableId` ancora impostato: un account AD con quell'ancora lo prenderebbe in carico |
| 4 - Farà soft-match con un account cloud | Un account AD non ancora sincronizzato e un account cloud con lo stesso UPN o SMTP primario: alla prima sincronizzazione verranno uniti |
| 5 - In conflitto con un account cloud | Stesso UPN / SMTP, ma l'account cloud ha già un altro `ImmutableId`: la sincronizzazione fallirà con un errore di attributo duplicato |
| 6 - Errore di provisioning in Entra | Errori che Entra segnala sull'oggetto (`proxyAddresses` duplicati, UPN...) |
| 7 - UPN in Entra ripiegato su onmicrosoft.com | L'attributo AD che la sincronizzazione usa come UPN (`userPrincipalName`, oppure `mail` con Alternate Login ID) ha un suffisso che non è un dominio verificato, quindi l'utente ha preso `@tenant.onmicrosoft.com`. Un UPN diverso da AD ma con un dominio verificato (es. `.local` in AD, `.it` in Entra) è una configurazione normale e non viene segnalato |
| 8 - Stessa ancora su più account AD | Oggetti copiati o ripristinati con lo stesso `ms-DS-ConsistencyGuid` |

Gli account sincronizzati da un altro dominio o foresta, o da fuori `SearchBase`, non vengono segnalati come mancanti. Se la maggior parte degli account sincronizzati non corrisponde per ancora (ancora personalizzata, es. `employeeID`), le categorie 1 e 2 vengono saltate e un avviso spiega il motivo.

## Configurazione (config.json)

| Chiave | Default | Descrizione |
|---|---|---|
| `Language` | di sistema | `en` o `it` |
| `Mode` | `FullHybrid` | `Cloud`, `FullHybrid`, `Hybrid`, `OnPrem` o `AD`. Modalità di partenza della finestra e modalità usata da `-Auto` |
| `TenantId` | *(vuoto)* | Dominio o ID del tenant. Vuoto vuol dire il tenant dell'account con cui accedi |
| `AuthMode` | `Interactive` | `Interactive` (browser) oppure `Certificate` (per `-Auto`) |
| `ClientId`, `CertificateThumbprint` | | Impostati da `tools\New-HybridScopeApp.ps1` |
| `IncludePerUserMfaState` | `true` | Legge lo stato MFA per-user legacy (una chiamata Graph per utente) |
| `DomainFqdn` | `%USERDNSDOMAIN%` | Dominio AD, es. `contoso.local`. Vuoto salta AD |
| `SearchBase` | *(vuoto)* | Limita AD a una OU, es. `OU=Azienda,DC=contoso,DC=local` |
| `UseWindowsCredentials`, `DomainUser` | `true`, vuoto | Ultimi valori usati nell'interfaccia. La password non viene mai salvata |
| `ExchangeServer` | *(vuoto)* | FQDN interno di un server Exchange. Vuoto salta Exchange |
| `ExchangeProtocol` | `http` | `http`, `https` o `local` |
| `ExchangeAuth` | `Kerberos` | `Kerberos`, `Negotiate` o `Basic` |
| `ExchangeSkipCertCheck` | `false` | Salta i controlli del certificato con `https` |
| `OutputFolder` | `Documenti\HybridScope` | Report, `history.csv` e log |
| `RetentionDays` | `90` | `-Auto` cancella report e log più vecchi (0 = mai) |
| `InactiveDays` | `90` | Soglia di inattività |
| `PasswordAgeDays` | `365` | Soglia password vecchia |
| `SyncStaleHours` | `3` | Tab SYNC: segnala la sincronizzazione se l'ultimo ciclo è più vecchio di così |
| `LicenseFilter` | `All` | `All` mostra tutti gli SKU (quelli sconosciuti con il nome tecnico); `MappedOnly` mostra solo quelli in `LicenseNames` |
| `LicenseNames` | SKU comuni | `"SKU_PART_NUMBER": "Nome leggibile"`. Aggiungi o rinomina liberamente |
| `GroupBy` | `TopLevelOU` | `TopLevelOU`, `ParentOU`, `None` (gruppo unico) oppure `Rules` (usa `GroupRules`). Selezionabile anche dalla finestra |
| `GroupRules` | `[]` | Usato solo con `GroupBy = Rules`. `[{ "Group": "DIPENDENTI", "Match": "OU=Dipendenti," }]`; vince il primo match sul DistinguishedName |
| `UnmappedGroup`, `CloudOnlyGroup` | `ALTRI`, `SOLO_CLOUD` | Nomi personalizzati per quei gruppi |
| `IncludeGuests` | `false` | Includi gli ospiti Entra |
| `ExportPreset`, `ExportColumns` | `Full`, `[]` | Scelta nel riquadro **Contenuto export**; `ExportColumns` sono le colonne di **Personalizzato** (nomi inglesi, es. `["Username", "AD_LastLogon", "M365_LastSignIn"]`) |
| `ExportGroupSheets`, `ExportOverview`, `ExportEnabledOnly` | `false` | Opzioni di export: un foglio per gruppo, includi PANORAMICA, solo account abilitati (valgono anche per `-Auto` con un preset) |
| `ExportPresets` | `{}` | I tuoi preset: `{ "Inattivi": ["Username", "AD_LastLogon", "M365_LastSignIn"] }`. Li salva la finestra di export |
| `AutoExportPreset` | `Full` | Contenuto del report di `-Auto`: `Full`, `Access`, `Licenses`, `Migration`, `MFA`, `Directory`, `Custom` o il nome di un tuo preset |

Esempio di `GroupRules`:
```json
"GroupRules": [
  { "Group": "DISABILITATI", "Match": "OU=Utenti Disabilitati," },
  { "Group": "INTERNI",      "Match": "OU=Interni," },
  { "Group": "ESTERNI",      "Match": "OU=Esterni," }
]
```
> Lascia la virgola finale in `Match`: evita match parziali come `OU=InterniOld`.

Per mostrare **solo alcune licenze** con un nome a scelta, imposta `"LicenseFilter": "MappedOnly"` e lascia in `LicenseNames` solo quelle che ti interessano:
```json
"LicenseFilter": "MappedOnly",
"LicenseNames": {
  "O365_BUSINESS_PREMIUM":    "MICROSOFT BUSINESS STANDARD",
  "O365_BUSINESS_ESSENTIALS": "MICROSOFT BUSINESS BASIC",
  "TEAMS_ESSENTIALS_AAD":     "TEAMS ESSENTIALS",
  "SPB":                      "BUSINESS PREMIUM"
}
```

## Permessi

HybridScope **legge e basta**. Bastano i privilegi minimi; per l'uso quotidiano non servono Domain Admin né Global Admin.

| Sistema | Minimo |
|---|---|
| Active Directory | Un qualsiasi utente di dominio |
| Exchange on-premises | Gruppo di ruoli **View-Only Organization Management** e `RemotePowerShellEnabled = $true` |
| Microsoft 365 | **Global Reader**, più un consenso amministratore una tantum ai permessi qui sotto |

Permessi Microsoft Graph (delegati nell'interfaccia, applicativi per `-Auto`):
`User.Read.All`, `AuditLog.Read.All`, `Reports.Read.All`, `UserAuthenticationMethod.Read.All`, `Organization.Read.All`, `Policy.Read.All`.

L'interfaccia chiede anche `GroupMember.Read.All` e `RoleManagement.Read.Directory`, in sola lettura, usati solo dalla scheda utente per mostrare i nomi di gruppi e ruoli. Se un amministratore non ha ancora dato il consenso, **Connetti 365** si collega comunque con i permessi qui sopra e la scheda mostra solo i gruppi AD; basta un consenso da amministratore per vedere anche gruppi e ruoli Entra.

## Riga di comando (senza interfaccia)

Ogni esecuzione senza finestra usa `config.json`; help completo con esempi: `Get-Help .\HybridScope.ps1 -Full`.

| Parametro | Cosa fa |
|---|---|
| `-Auto` | Report completo per la `Mode` configurata, scritto in `OutputFolder` (più `history.csv` e `HybridScope_LATEST.xlsx`) |
| `-Only <dataset>` | Un solo insieme di dati: `MFA`, `Licenses`, `Mailbox365`, `OneDrive`, `MailboxOnPrem` o `Sync` (AD + Microsoft 365). Vengono contattate solo le fonti che servono |
| `-PassThru` | Restituisce oggetti invece di scrivere file: le righe di `-Only`, oppure una riga unita per utente. I valori sono sempre in inglese |
| `-Mode <modalità>` | Sostituisce `Mode` solo per questa esecuzione (`Cloud`, `FullHybrid`, `Hybrid`, `OnPrem`, `AD`) |
| `-ExportPreset <nome>` | Contenuto del report: `Full` (default), `Access`, `Licenses`, `Migration`, `MFA`, `Directory`, `Custom` o un tuo preset. Sostituisce `AutoExportPreset` |
| `-Columns <lista>` | Solo queste colonne della lista utenti, in questo ordine, es. `Username,AD_LastLogon,M365_LastSignIn` (nomi inglesi o intestazioni italiane) |
| `-ConfigPath <file>` | Usa un altro file di configurazione, per esempio uno per tenant |

Accesso a Microsoft 365: con `AuthMode = Interactive` (il default) si apre una volta l'accesso nel browser; con `Certificate` nessuna richiesta. Se nella stessa sessione hai già eseguito `Connect-MgGraph`, quella connessione viene riusata e lasciata aperta.

```powershell
# solo lo stato MFA, in Excel nella OutputFolder (AD ed Exchange non vengono toccati)
.\HybridScope.ps1 -Only MFA

# utenti senza MFA, direttamente in CSV
.\HybridScope.ps1 -Only MFA -PassThru | Where-Object MFA_Registered -eq 'No' | Export-Csv .\no-mfa.csv -NoTypeInformation

# cosa andrà storto alla prossima sincronizzazione AD -> Entra ID
.\HybridScope.ps1 -Only Sync -PassThru | Where-Object Category -notlike '0 *' | Format-Table Category, Username, UPN, Details -Wrap

# solo username, ultimo logon AD e ultimo accesso Microsoft 365
.\HybridScope.ps1 -Auto -Columns Username,AD_LastLogon,M365_LastSignIn

# report MFA pronto
.\HybridScope.ps1 -Auto -ExportPreset MFA

# report completo, solo Microsoft 365, senza modificare config.json
.\HybridScope.ps1 -Auto -Mode Cloud

# dati uniti per utente come oggetti: mailbox on-prem da migrare, dalla più grande
.\HybridScope.ps1 -PassThru | Where-Object Location -eq 'ON-PREM' | Sort-Object MailboxOnPrem_MB -Descending
```

Exit code: `0` = riuscito, `1` = errore (dettagli in `OutputFolder\logs`).

## Report pianificati (-Auto)

1. **Crea l'app registration** con certificato. Va fatto una volta, da Global Admin, **con l'account che eseguirà il task**:
   ```powershell
   .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com
   ```
   Crea il certificato, l'app con permessi in sola lettura e il consenso amministratore, e aggiorna `config.json` (`AuthMode = Certificate`).
2. **Registra il task giornaliero**, da PowerShell come amministratore:
   ```powershell
   .\tools\Register-HybridScopeTask.ps1 -RunAs CONTOSO\svc-hybridscope -At 06:40
   ```
3. Ogni esecuzione scrive in `OutputFolder`:
   - `HybridScope_<data>.xlsx` (con un preset: `HybridScope_<preset>_<data>.xlsx`, vedi `AutoExportPreset`)
   - `HybridScope_LATEST.xlsx`, sempre l'ultimo, comodo da condividere
   - `history.csv`, mostrato anche nel tab STORICO
   - `logs\`

Esecuzione manuale: `powershell -File .\HybridScope.ps1 -Auto`

## Risoluzione problemi

<details>
<summary><b>Collegato a Microsoft 365 con l'account sbagliato</b></summary>

Premi **Cambia account** per disconnetterti e cancellare l'accesso salvato, poi di nuovo **Connetti 365**. Se il browser sceglie da solo l'account sbagliato, nella pagina di accesso Microsoft scegli *Usa un altro account*, oppure esci prima da quell'account nel browser.
</details>

<details>
<summary><b>"WorkgroupCannotUseKerberos" / "Impossibile utilizzare l'autenticazione Kerberos con credenziali implicite"</b></summary>

Il PC non è nel dominio. Togli **Usa credenziali Windows** e inserisci `utente@dominio.local`. Il PC deve risolvere il dominio tramite DNS interno o VPN: `nslookup -type=SRV _kerberos._tcp.contoso.local`.
</details>

<details>
<summary><b>"Nome utente o password non corretta" verso Exchange</b></summary>

Controlla il sAMAccountName esatto, per esempio `m.rossi` invece di `mrossi`, e usa il formato `utente@dominio.local`. Prova fuori da HybridScope:

```powershell
$cred = Get-Credential utente@contoso.local
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch01.contoso.local/PowerShell/ -Authentication Kerberos -Credential $cred
Invoke-Command -Session $s -ScriptBlock { Get-ExchangeServer }
```
</details>

<details>
<summary><b>Il browser chiede un certificato su https://server/powershell</b></summary>

È normale: `/PowerShell` è un endpoint WinRM, non una pagina web. Usa `http` + `Kerberos` con il nome interno del server.
</details>

<details>
<summary><b>Le dimensioni di mailbox 365 / OneDrive sono vuote</b></summary>

I report di utilizzo sono anonimizzati. Nell'interfaccia di amministrazione di Microsoft 365 vai in **Impostazioni > Impostazioni organizzazione > Report** e togli la spunta a *Visualizza nomi utente, gruppi e siti nascosti in tutti i report*. I report possono anche essere in ritardo di 24-48 ore.
</details>

<details>
<summary><b>Modulo ActiveDirectory: errore 0x800f0954</b></summary>

Il WSUS blocca le funzionalità su richiesta. Installa *RSAT: Strumenti di Servizi di dominio Active Directory* da **Impostazioni > Sistema > Funzionalità facoltative**, oppure chiedi all'IT di consentirlo. HybridScope non modifica le impostazioni di Windows Update.
</details>

<details>
<summary><b>L'antivirus segnala l'exe</b></summary>

Gli eseguibili creati con PS2EXE (`tools\Build-Exe.ps1`) vengono spesso segnalati per errore (falsi positivi). Usa `Start-HybridScope.cmd`, che avvia lo script in chiaro con Windows PowerShell, oppure firma l'exe con il certificato di code signing aziendale.
</details>

<details>
<summary><b>"l'esecuzione di script è disabilitata" / "non è firmato digitalmente"</b></summary>

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Get-ChildItem C:\HybridScope -Recurse | Unblock-File
```
</details>

## FAQ

**Modifica qualcosa?** No. Tutte le chiamate sono in sola lettura. L'unica eccezione è l'installazione dei moduli PowerShell mancanti, e solo se accetti.

**Dove vengono inviati i miei dati?** Da nessuna parte. HybridScope comunica solo con il tuo AD, il tuo server Exchange e Microsoft Graph. I report restano sul tuo disco. L'unica altra connessione è **Aiuto > Cerca aggiornamenti**, e solo quando lo premi: legge il numero dell'ultima release dalle API pubbliche di GitHub (`api.github.com`) e non invia nulla su di te o sul tuo tenant.

**Funziona senza Microsoft 365, con solo Exchange 2016/2019 e AD?** Sì: usa la modalità **On-Prem**.

**Solo Microsoft 365, oppure Microsoft 365 + AD senza Exchange?** Sì: usa la modalità **Full Cloud** o **Hybrid**.

**Solo Active Directory?** Sì: usa la modalità **AD**.

**Posso estrarre solo lo stato MFA, senza interfaccia?** Sì: `.\HybridScope.ps1 -Only MFA` (Excel) oppure `-Only MFA -PassThru` (oggetti). Vedi [Riga di comando](#riga-di-comando-senza-interfaccia).

## Roadmap

- [x] Scheda utente: doppio clic su un utente per vedere insieme AD, mailbox, OneDrive, MFA, licenze, gruppi e ruoli
- [x] Controllo della sincronizzazione: ImmutableId / ancora di origine, soft-match e conflitti prima della prossima sincronizzazione
- [ ] Mailbox condivise / sale e liste di distribuzione: condivise con accesso abilitato, licenze inutili, deleghe, liste vuote o senza proprietario
- [ ] Riepilogo HTML dopo `-Auto`: panoramica, lista da attenzionare e andamento, da mandare via mail

## Sostieni il progetto

HybridScope è gratuito, open source e resterà così. Se HybridScope (o un altro dei miei progetti) ti fa risparmiare tempo, puoi lasciare una **donazione una tantum**: nessun abbonamento, nessun rinnovo automatico.

| | Importo | Per chi |
|---|---|---|
| ☕ **Coffee** | $5 | Un grazie veloce |
| 🍕 **Lunch** | $10 | Lo usi spesso e vuoi che resti gratuito, in sola lettura e open source |
| 🛠️ **Sysadmin** | $25 | Ti ha fatto risparmiare ore, in una migrazione, un audit o una pulizia delle licenze |
| 🏢 **Company** | $100 | Aziende e MSP che lo usano in produzione. Se vuoi, il nome della tua azienda compare nei [Ringraziamenti](#ringraziamenti) |

[![Sponsor su GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Offrimi un caffè](https://img.shields.io/badge/Offrimi%20un%20caff%C3%A8-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/aascione956)
[![Metti una stella](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Stella)](https://github.com/Thonno/hybridscope/stargazers)

Non vuoi donare? Aiuta lo stesso: una ⭐, una issue con un bug o un'idea, o consigliare HybridScope a un collega.

Issue e pull request sono benvenute. Per modifiche importanti apri prima una issue.

## Ringraziamenti

Grazie a chi sostiene HybridScope.

**Aziende e MSP**

<!-- Aggiungi qui le aziende del livello Company che hanno chiesto di comparire, una per riga, es.:
- [Contoso Srl](https://www.contoso.it)
-->
*Ancora nessuno: la tua azienda potrebbe essere la prima.*

## Licenza

[MIT](LICENSE). HybridScope è un progetto indipendente e non è affiliato, approvato o sponsorizzato da Microsoft. Microsoft 365, Exchange, Entra ID e Active Directory sono marchi del gruppo Microsoft.
