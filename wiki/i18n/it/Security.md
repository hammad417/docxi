<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Sicurezza di docxi: segnalare problemi e gestire input non attendibili

Questa pagina spiega come segnalare un problema di sicurezza e quali aspetti di sicurezza conoscere quando la tua applicazione passa a docxi dati forniti dagli utenti.

## <a id="reporting-a-vulnerability"></a>Segnalare una vulnerabilità

Per favore, non segnalare problemi di sicurezza in issue pubbliche. Scrivi invece agli indirizzi di manutenzione riportati in [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Indica la versione di docxi e un modo rapido per riprodurre il problema.

## <a id="upgrade-from-00x"></a>Aggiorna dalla 0.0.x

docxi 0.0.x passa le stringhe dei percorsi delle immagini a `Kernel#open`. `Kernel#open` esegue come comando di shell una stringa che inizia con `|`, quindi chi controlla il percorso di un'immagine potrebbe eseguire comandi sul tuo server. La 0.1.0 apre i file con `File.open` e scarica gli URL con `URI#open`. Nessuno dei due esegue comandi.

## <a id="untrusted-input"></a>Input non attendibili

**Il testo si può passare in sicurezza.** Il testo, il contenuto delle tabelle e le altre stringhe vengono scritti come testo XML o valori di attributo, a cui nokogiri applica l'escape. L'input degli utenti non può aggiungere elementi al documento.

**Le sorgenti delle immagini richiedono attenzione.** `add_media` accetta:

| Sorgente | Cosa fa docxi |
|---|---|
| Un IO che hai aperto tu | Lo legge |
| Un percorso di file | Legge quel file dal disco del tuo server |
| Un URL `http`, `https` o `ftp` | Lo scarica dal tuo server |

Se un utente può scegliere il percorso, può far incorporare nel documento qualsiasi file leggibile del tuo server, come la configurazione o le chiavi. Se può scegliere l'URL, il tuo server lo richiederà, e la richiesta può raggiungere indirizzi interni (falsificazione delle richieste lato server, SSRF). Passa un IO che hai aperto tu, oppure consenti solo percorsi e host noti.

**I collegamenti ipertestuali** puntano dove indica l'URL. Controlla gli URL forniti dagli utenti prima di passarli ad `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Come sono protette le versioni

- Le versioni vengono pubblicate da GitHub Actions con la pubblicazione attendibile (trusted publishing) di RubyGems, quindi nel repository non è conservata alcuna chiave API.
- I metadati della gem impostano `rubygems_mfa_required`, quindi le pubblicazioni fatte con una chiave API personale richiedono l'autenticazione a più fattori.

## <a id="related-pages"></a>Pagine correlate

- [Buone pratiche](Best-Practices)
- [Deployment](Deployment)
- [Contribuire](Contributing)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
