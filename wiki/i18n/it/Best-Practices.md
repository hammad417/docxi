<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Buone pratiche con docxi: evitare le sorprese più comuni

Queste abitudini derivano direttamente da come funziona docxi, e la maggior parte previene un problema elencato in [Risoluzione dei problemi](Troubleshooting). Sono raggruppate per argomento. I consigli sulle impostazioni di pagina valgono per tutti, mentre quelli su sicurezza e prestazioni contano soprattutto per le applicazioni in produzione.

## <a id="page-setup"></a>Impostazioni di pagina

- **Passa i quattro margini insieme.** L'hash che passi a `Docxi::Document.new` sostituisce i valori predefiniti, e qualsiasi valore tra `top`, `right`, `bottom` e `left` che ometti viene scritto vuoto.
- **Imposta l'intestazione e il piè di pagina della prima pagina.** docxi attiva sempre una prima pagina distinta. Se la pagina 1 ha bisogno di un'intestazione o di un piè di pagina, chiama anche `add_first_header` e `add_first_footer`.

## <a id="text"></a>Testo

- **Usa `space: true` per gli spazi iniziali o finali.** Altrimenti Word li elimina e le parole di due frammenti si attaccano.
- **Metti la formattazione comune sul paragrafo.** Le opzioni passate a `p`, come `size:` o `font:`, si applicano a ogni frammento al suo interno, così c'è meno da ripetere.
- **Usa `p.br` per le interruzioni di riga e `document.br(page: true)` per le interruzioni di pagina.** `document.br` senza `page: true` produce XML non valido.

## <a id="tables"></a>Tabelle

- **Imposta i bordi sulle righe o sulle celle, non sulla tabella.** Nella 0.1.0 i `borders:` a livello di tabella vengono applicati a ogni cella invece che al contorno della tabella.
- **Indica `columns_width` per ogni colonna.** Le celle ricavano da lì la propria larghezza, così le tabelle mantengono il layout che ti aspetti.

## <a id="images"></a>Immagini

- **Usa file PNG.** Ogni immagine viene memorizzata come parte PNG.
- **Dimensiona le immagini in linea con `style:`**, per esempio `style: { width: "120pt", height: "60pt" }`. È l'unica impostazione delle dimensioni delle immagini che nella 0.1.0 funziona come ci si aspetta.
- **Passa un percorso anziché un `File` aperto se il documento potrebbe essere generato più di una volta.** Le immagini IO vengono lette durante `render`, e una seconda generazione ottiene un'immagine vuota.
- **Immagini nelle intestazioni, nelle celle delle tabelle e immagini posizionate fino alla 0.2.0.** Larghezza e altezza vengono entrambe calcolate come `width * height * 14.92` EMU, quindi risultano quadrate. Per ottenere un'immagine quadrata di `S` pollici di lato, passa `width` e `height` pari a `Math.sqrt(S * 914400 / 14.92).round`, cioè circa 248 per un pollice.

## <a id="security"></a>Sicurezza

- **Non passare stringhe non attendibili ad `add_media`.** Un percorso fa leggere a docxi quel file dal tuo server, e un URL lo fa scaricare al tuo server. Passa un IO che hai aperto tu, oppure controlla prima il valore. Consulta [Sicurezza](Security).
- **Il testo viene sottoposto a escape per te.** Le stringhe passate a `text` vengono scritte come testo XML, quindi l'input degli utenti non può modificare la struttura del documento.

## <a id="performance-and-reliability"></a>Prestazioni e affidabilità

- **Genera i documenti grandi in un job in background.** L'intero documento viene costruito in memoria quando chiami `render`. Consulta [Integrazioni](Integrations).
- **Genera una volta, poi riusa i byte.** Conserva `word.render.string` se devi inviare o salvare il documento più di una volta.
- **Fissa una versione compatibile** con `gem "docxi", "~> 0.1"`.
- **Testa ciò che generi.** Un `.docx` è uno zip di file XML, quindi i test possono decomprimerlo e controllare l'XML. I test di docxi funzionano proprio così. Consulta [Architettura](Architecture).

## <a id="related-pages"></a>Pagine correlate

- [Funzionalità](Features)
- [Configurazione](Configuration)
- [Risoluzione dei problemi](Troubleshooting)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
