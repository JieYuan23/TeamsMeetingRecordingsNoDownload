# Set Teams Meeting Recordings as no download
La soluzione nasce per la necessità di inibire, in seguito alla registrazione di un public channel meeting, il download della registrazione effettuata.

By Design
La registrazione viene salvata sullo storage SharePoint dedicato al Teams.
I partecipanti al meeting hanno i permessi per visualizzare/modificare la registrazione ed effettuare il download del file.

Per modificare il comportamento By Design del prodotto, e la normale assegnazione dei permessi, è stato implementato uno script PowerShell.
Il codice effettua la modifica del permessi sulla cartella contenente la registrazione assegnando delle ACL specifiche per inibire la possibilità di download dei files che contiene.

Lo sviluppo della procedura ha richiesto l’utilizzo dei seguenti moduli PS:
MicrosoftTeams
SharePointPnPPowerShellOnline
Microsoft Graph API
I moduli, nel caso non dovessero essere presenti sul client/server, saranno automaticamente installati dalla procedura.
Preconditions
La login utilizzata a parametro deve essere SharePoint Admin e Teams Admin

PS. è attualmente in sviluppo la versione automatizzata che schedula l'esecuzione, si basa su Azure Functions.

**PowerShell version**
Nome script: DisableTeamsRecordingDownloadFromSPO.ps1
Prerequisiti
L’utente (fornito come parametro) che verrà utilizzato dallo script deve essere amministratore di Teams.
Lo script utilizza i moduli PowerShell “MicrosoftTeams” e “SharePointPnPPowerShellOnline”, qualora non fossero già installati provvede alla relative installazione. Importante: la versione del modulo “SharePointPnPPowerShellOnline“ con cui è stato testato lo script è la “3.25.2009.1”, non utilizzare la versione “3.26.2010.0” in quanto introduce dei problemi di autenticazione.
Parametri
- User (identificativo dell’utente che lo script utilizzerà per effettuare le operazioni, deve essere amministratore di Teams).
- MFAUser (switch da utilizzare solamente nel caso in cui l’utente abbia la MFA abilitata).
- TeamDisplayName (per eseguire lo script solamente per uno specifico team, se non utilizzato lo script agirà su tutti i Team presenti nel tenant).
- CreateRecordingsFolder (switch da utilizzare per creare preventivamente le cartelle “Recordings” all’interno delle cartelle dei canali e cambiare i permessi per non consentire il download delle registrazioni).
**Comportamento**
Lo script cambia i permessi dei membri del Team sulle cartelle “Recordings” per impedire loro di poter scaricare le registrazioni.
Esempio
.\DisableTeamsRecordingDownloadFromSPOCurrent.ps1 -User andrear@M365EDU702432.onmicrosoft.com -MFAUser -TeamDisplayName TeamDiProva -CreateRecordingsFolder
Output
Lo script mostra a video le attività che vengono eseguite, le stesse vengono catturate in un file di testo (.txt).
Viene inoltre generato un file di report (.csv) che riporta il dettaglio dei permessi assegnati ai gruppi SharePoint visitatori di / membri di / proprietari di su ogni cartella contenente le registrazioni (“Recordings”).
Nota: entrambi i file vengono salvati nella cartella dalla quale viene eseguito lo script.
