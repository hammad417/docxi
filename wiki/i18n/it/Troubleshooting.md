<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Risoluzione dei problemi di docxi: errori di installazione, errori in esecuzione e documenti difettosi

Cerca il tuo sintomo qui sotto. La maggior parte dei problemi nelle applicazioni datate deriva da docxi 0.0.x, e l'aggiornamento alla 0.1.0 li risolve senza modifiche al codice.

## <a id="installation"></a>Installazione

**Bundler non trova versioni compatibili di `nokogiri` in un'applicazione Rails 6+.**
docxi 0.0.6 blocca `nokogiri < 1.7`, mentre ogni versione di Rails dalla 6.0 in poi richiede nokogiri 1.8.5 o successivo. Usa `gem "docxi", "~> 0.1"`.

**Bundler non trova versioni compatibili di `rubyzip`.**
docxi 0.0.6 blocca `rubyzip ~> 0.9.9`, che va in conflitto con qualsiasi gem che richieda un rubyzip più recente. docxi 0.1.0 accetta rubyzip dalla 1.3 alla 3.x.

**Bundler continua a installare la 0.0.6.**
O il Gemfile fissa docxi a una 0.0.x, per esempio con `"~> 0.0.6"`, oppure Ruby è precedente alla 2.5, richiesta dalla 0.1.0. Cambia la riga del Gemfile in `gem "docxi", "~> 0.1"` e controlla `ruby -v`.

**La compilazione di nokogiri non riesce.**
Nessuna gem precompilata di nokogiri corrisponde alla tua piattaforma, quindi viene compilata dai sorgenti. Installa un compilatore C e segui la [guida all'installazione di nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Errori durante la creazione di un documento

**`NoMethodError: undefined method 'present?'` fuori da Rails.**
docxi 0.0.6 presupponeva che ActiveSupport fosse caricato. La 0.1.0 non ne ha bisogno.

**`NameError` chiamando `tab`, `frame` o il `page_numbers` di un'intestazione.**
Questi metodi non funzionavano nella 0.0.6 e funzionano nella 0.1.0, compreso `document.frame("text")`.

**`Errno::ENOENT` aggiungendo un'immagine.**
Il percorso non esiste rispetto alla directory in cui gira il tuo programma. Con Ruby 3 e docxi 0.0.6 anche gli URL delle immagini falliscono in questo modo, perché la 0.0.6 usava `Kernel#open`. La 0.1.0 scarica gli URL con `URI#open`.

## <a id="documents-that-dont-look-right"></a>Documenti che non appaiono come dovrebbero

**Google Docs o LibreOffice non aprono il file.**
docxi 0.0.6 con rubyzip 3 scrive record Zip64 in ogni voce, e questi programmi li rifiutano. La 0.1.0 scrive gli archivi senza di essi.

**La prima pagina non ha intestazione né piè di pagina.**
docxi assegna alla prima pagina un'intestazione e un piè di pagina propri. Impostali con `add_first_header` e `add_first_footer`. Consulta [Configurazione](Configuration#first-page).

**Manca un margine, oppure il layout della pagina sembra sbagliato.**
Quando imposti i margini, passa `top`, `right`, `bottom` e `left` insieme. Quelli che ometti vengono scritti vuoti.

**Gli spazi tra i frammenti di testo scompaiono.**
Word elimina gli spazi all'inizio o alla fine di un frammento. Aggiungi `space: true` a quella chiamata a `text`.

**Le immagini nelle intestazioni, nelle celle delle tabelle o con `pH`/`pV` sono minuscole o quadrate.**
Problema noto nella 0.1.0: le loro dimensioni vengono calcolate come `width * height * 14.92` EMU per entrambe le dimensioni. Usa immagini in linea dimensionate con `style:`, oppure consulta la soluzione alternativa in [Buone pratiche](Best-Practices#images).

**Un'immagine manca la seconda volta che si genera lo stesso documento.**
Le immagini aggiunte da un `File` aperto o da un altro IO vengono lette durante `render`, e una seconda generazione legge un flusso vuoto. Passa invece un percorso, oppure genera il documento una volta sola.

**I bordi impostati su `table` compaiono attorno a ogni cella.**
Problema noto: i `borders:` a livello di tabella vengono applicati a ogni cella invece che al contorno della tabella. Per disegnare solo il contorno, imposta `borders:` sulle celle del perimetro, solo per i loro lati esterni.

**I numeri di pagina nelle celle delle tabelle iniziano con «GlobalOptions».**
Problema noto di `TableCell#page_numbers`. Metti invece i numeri di pagina in un piè di pagina o in un'intestazione.

**Un'immagine JPEG o GIF appare sbagliata o non compare.**
Ogni immagine viene memorizzata come parte PNG con un tipo di contenuto PNG, qualunque sia il suo formato reale, quindi alcuni programmi potrebbero non mostrare gli altri formati. Converti le immagini in PNG prima di aggiungerle.

## <a id="inspecting-a-generated-file"></a>Ispezionare un file generato

Un `.docx` è un archivio zip, quindi puoi guardarci dentro:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Se il problema persiste, [apri una issue](https://github.com/hammad417/docxi/issues) con un breve script che lo riproduca e le versioni di Ruby, docxi, nokogiri e rubyzip.

## <a id="related-pages"></a>Pagine correlate

- [Domande frequenti](FAQ)
- [Buone pratiche](Best-Practices)
- [Installazione](Installation)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
