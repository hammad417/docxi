<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Funzionalità di docxi: paragrafi, testo, tabelle, immagini, intestazioni e piè di pagina

Questa pagina elenca ogni elemento che docxi può aggiungere a un documento Word e le opzioni accettate da ciascuno. Gli esempi partono da questa configurazione:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Le opzioni che non ti servono si possono omettere. I colori sono valori esadecimali senza `#`, per esempio `"8B0000"`. Per le unità di misura e le impostazioni di pagina, consulta [Configurazione](Configuration).

## <a id="paragraphs"></a>Paragrafi

`document.p(options) { |p| ... }` aggiunge un paragrafo.

| Opzione | Valori | Effetto |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Allineamento. `"both"` giustifica il testo |
| `fill` | Colore esadecimale | Ombreggiatura dietro il paragrafo e il suo testo |
| `left`, `right` | Ventesimi di punto | Rientro. Imposta entrambi, perché uno solo viene ignorato |
| `ul` | `true` | Trasforma il paragrafo in una voce di elenco puntato |
| `bottom` | `true` | Disegna un bordo sotto il paragrafo |

Anche le opzioni del testo come `size:`, `bold:` o `color:` si possono passare a `p`. Diventano i valori predefiniti di ogni frammento di testo del paragrafo.

All'interno del blocco puoi chiamare `text`, `br`, `tab`, `image`, `hyperlink` e `frame`. `document.text(string, options)` è una scorciatoia per un paragrafo con un solo frammento.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Testo

`p.text(string, options)` aggiunge un frammento di testo (run).

| Opzione | Valori | Effetto |
|---|---|---|
| `bold` | `true` | Grassetto |
| `italic` | `true` | Corsivo |
| `underline` | `true` | Sottolineatura singola |
| `color` | Colore esadecimale | Colore del testo |
| `size` | Punti | Dimensione del carattere |
| `font` | Nome del carattere | Carattere, per esempio `"Georgia"` |
| `fill` | Colore esadecimale | Ombreggiatura dietro il testo |
| `space` | `true` | Conserva gli spazi all'inizio e alla fine del testo |
| `br` | `true` | Aggiunge un'interruzione di riga dopo il testo |

Una stringa `nil` viene saltata.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Interruzioni di riga, interruzioni di pagina e tabulazioni

- `p.br`, oppure `p.br(times: 3)`, aggiunge interruzioni di riga dentro un paragrafo.
- `document.br(page: true)` inizia una nuova pagina.
- `p.tab`, oppure `p.tab(times: 2)`, aggiunge caratteri di tabulazione dentro un paragrafo.
- `document.tab(page: true)` aggiunge un paragrafo che contiene una sola tabulazione.

Non chiamare `document.br` o `document.tab` senza `page: true`. Inseriscono un frammento di testo direttamente nel corpo del documento, e questo non è WordprocessingML valido.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tabelle

`document.table(options)` riceve un blocco, al cui interno `table.tr(options)` aggiunge righe e `tr.tc(options)` aggiunge celle.

**Opzioni della tabella**

| Opzione | Valori | Effetto |
|---|---|---|
| `columns_width` | Array di numeri | Larghezze delle colonne, moltiplicate per 14.92 in ventesimi di punto. Sono anche le larghezze predefinite delle celle |
| `height` | Ventesimi di punto | Altezza minima di ogni riga |
| `iwidth` | Ventesimi di punto | Rientro della tabella rispetto al margine sinistro |
| `style` | ID dello stile | Stile della tabella. docxi definisce `TableNormal` e `TableCustom` |
| `width` | Numero | Larghezza preferita della tabella, scritta con il tipo `auto`. Senza `columns_width`, anche ogni cella riceve questa larghezza, moltiplicata per 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Allineamento della tabella. Il testo aggiunto con `tc.text` viene allineato allo stesso modo |
| `padding` | `true` | Aggiunge una spaziatura interna fissa nelle celle |
| `borders` | Hash | Problema noto: viene applicata a ogni cella invece che al contorno della tabella |

**Opzioni della riga**: `height:` imposta l'altezza minima della riga. Qualsiasi opzione di cella passata a `tr` si applica a ogni cella della riga. Le opzioni della tabella vengono passate allo stesso modo a ogni riga e cella.

**Opzioni della cella**

| Opzione | Valori | Effetto |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Spessore del bordo in ottavi di punto. `nil` significa nessun bordo su quel lato |
| `fill` | Colore esadecimale | Ombreggiatura della cella. Le celle senza riempimento sono bianche (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Allineamento verticale |
| `colspan` | Intero | Numero di colonne occupate dalla cella |
| `rowspan` | `true` | Avvia un'unione verticale |
| `merged` | `true` | Prosegue l'unione verticale della cella sopra. Il suo contenuto viene ignorato |
| `width` | Numero | Larghezza della cella, moltiplicata per 14.92 |

**Contenuto della cella**

- `tc.text(string, options)`: le opzioni della cella e della riga, come `bold:` o `align:`, si applicano anche al testo.
- `tc.p(options) { |p| ... }`: un paragrafo che usa solo le proprie opzioni.
- `tc.image(media, width:, height:)`: un'immagine. Per il problema delle dimensioni vedi [Immagini](#images).
- `tc.page_numbers`: un campo numero di pagina. Problema noto: è preceduto dal testo «GlobalOptions».

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Immagini

`media = document.add_media(source)` registra un'immagine. `source` può essere:

- un oggetto IO, come un `File` o uno `StringIO`
- un percorso di file
- un URL `http`, `https` o `ftp`

L'immagine viene letta quando chiami `render` ed è sempre memorizzata come parte PNG, quindi i file PNG sono la scelta sicura.

**Immagini in linea**: `p.image(media, options)` colloca l'immagine nella riga di testo.

| Opzione | Valori | Effetto |
|---|---|---|
| `style` | Hash di proprietà CSS | Dimensioni, per esempio `{ width: "120pt", height: "60pt" }` |
| `title` | Stringa | Titolo dell'immagine |
| `fill_color` | Colore VML come `"#00FF00"` | Riempimento dietro l'immagine |
| `opacity` | Stringa come `"50%"` | Opacità del riempimento |
| `wrap` | `true` | Disattiva la disposizione del testo attorno all'immagine |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Immagini posizionate**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` ancora l'immagine con uno scostamento dalla colonna (`pH`) e dal paragrafo (`pV`), in EMU. `pH`, `pV` e `wrap` vanno indicati tutti, altrimenti l'immagine viene collocata in linea.

**Immagini nell'intestazione**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` disattiva la disposizione del testo. Qualsiasi altro valore dispone il testo attorno all'immagine con un ingombro quadrato.

**Problema noto nella 0.1.0**: le immagini posizionate, quelle nelle intestazioni e quelle nelle celle delle tabelle calcolano sia la larghezza sia l'altezza come `width * height * 14.92` EMU. Risultano sempre quadrate e di solito minuscole. Le immagini in linea dimensionate con `style:` non ne sono interessate. Consulta [Buone pratiche](Best-Practices#images) per una soluzione alternativa.

## <a id="headers-and-footers"></a>Intestazioni e piè di pagina

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Un'intestazione può contenere `text`, `br`, `table`, `image` e `page_numbers`. Un piè di pagina può contenere `text`, `br`, `table` e `page_numbers`.
- Le opzioni passate a `Header.new` o `Footer.new` si applicano al loro testo.
- `add_header` e `add_footer` valgono dalla pagina 2 in poi. La pagina 1 usa sempre `add_first_header` e `add_first_footer`. Consulta [Configurazione](Configuration#first-page).

## <a id="page-numbers"></a>Numeri di pagina

`page_numbers(name: "Page", align: "center")` scrive il testo di `name` seguito da un campo PAGE. Il numero è in Arial 10 pt, grigio scuro. Il valore predefinito di `align` è `"right"`.

Problema noto: nel codice di campo manca la barra rovesciata di `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Sommario

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Con un numero di pagina, la voce riceve un punto di tabulazione allineato a destra in `pos` (in ventesimi di punto; 9360 corrisponde a 6,5 pollici), riempito con il carattere di `leader`. I valori del carattere di riempimento provengono da WordprocessingML, come `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` e `"none"`.
- Senza numero di pagina, la voce è il testo dopo una tabulazione.
- I numeri di pagina sono quelli che passi. docxi non crea un campo TOC di Word e non calcola le pagine.

## <a id="hyperlinks"></a>Collegamenti ipertestuali

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` restituisce un ID di relazione, e `hyperlink(text, url, id)` inserisce il collegamento. La destinazione viene dalla relazione, quindi passa l'ID restituito per l'URL che ti interessa. Sono supportati solo i collegamenti esterni.

## <a id="frames-text-boxes"></a>Riquadri (caselle di testo)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- I riquadri sono caselle di testo VML con un riempimento fisso grigio chiaro semitrasparente e un contorno nero. Il testo si dispone sopra e sotto di essi.
- Dentro un riquadro puoi chiamare `text`, `br`, `tab`, `image` e `frame`.
- `spacing: true` aggiunge spazio prima del testo.

## <a id="output"></a>Output

`word.render` restituisce uno `StringIO` con il pacchetto `.docx` completo. L'archivio viene scritto senza record Zip64, così Google Docs e le versioni meno recenti di LibreOffice riescono ad aprirlo. [Architettura](Architecture) elenca ogni parte del pacchetto.

## <a id="related-pages"></a>Pagine correlate

- [Esempi d'uso](Usage-Examples)
- [Riferimento API](API-Reference)
- [Roadmap](Roadmap), per i problemi noti previsti per la 0.2.0

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
