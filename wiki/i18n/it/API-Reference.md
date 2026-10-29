<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Riferimento API di docxi: classi e metodi

docxi è una libreria Ruby, quindi questa pagina elenca la sua API Ruby. Non esiste un'API HTTP. Le opzioni di ciascun metodo sono descritte in [Funzionalità](Features), e le classi sono spiegate in [Architettura](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Metodo | Restituisce | Note |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` sono i margini della pagina. Consulta [Configurazione](Configuration) |
| `#document` | `Docxi::Word::Document` | Il corpo a cui aggiungi i contenuti |
| `#render` | `StringIO` | Il pacchetto `.docx` completo. `.string` restituisce i byte |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Metodo | Restituisce | Aggiunge |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Un paragrafo |
| `text(string, options = {})` | Paragraph | Un paragrafo con un solo frammento di testo |
| `br(page: true)` | Break | Un'interruzione di pagina |
| `tab(page: true)` | Tab | Un paragrafo con una sola tabulazione |
| `table(options = {}) { \|table\| }` | Table | Una tabella |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Voci del sommario |
| `frame(text, options = {})` | Paragraph | Un paragrafo con una casella di testo |
| `add_media(io_path_or_url, options = {})` | Media | Un'immagine da collocare con `image` |
| `add_hyperlink(url, options = {})` | String | Un ID di relazione per `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Intestazione dalla pagina 2 in poi e per la pagina 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Piè di pagina dalla pagina 2 in poi e per la pagina 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Paragrafo (`Docxi::Word::Contents::Paragraph`)

| Metodo | Aggiunge |
|---|---|
| `text(string, options = {})` | Un frammento di testo. Le opzioni del paragrafo sono i suoi valori predefiniti |
| `br(times: 1)` | Interruzioni di riga |
| `tab(times: 1)` | Caratteri di tabulazione |
| `image(media, options = {})` | Un'immagine in linea o posizionata |
| `hyperlink(text, url, id)` | Un collegamento che usa l'`id` di relazione restituito da `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Una casella di testo |

## <a id="tables"></a>Tabelle

| Metodo | Aggiunge |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Una riga. Le opzioni della tabella sono i suoi valori predefiniti |
| `TableRow#tc(options = {}) { \|tc\| }` | Una cella. Le opzioni della riga sono i suoi valori predefiniti |
| `TableCell#text(string, options = {})` | Un paragrafo. Le opzioni della cella sono i suoi valori predefiniti |
| `TableCell#p(options = {}) { \|p\| }` | Un paragrafo con le proprie opzioni |
| `TableCell#image(media, width:, height:)` | Un'immagine |
| `TableCell#page_numbers` | Un campo numero di pagina |

## <a id="table-of-contents"></a>Sommario

| Metodo | Aggiunge |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Una voce del sommario |

## <a id="headers-and-footers"></a>Intestazioni e piè di pagina

| Classe | Metodi |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Riquadri

`Docxi::Word::Contents::Frame` accetta `text`, `br`, `tab`, `image` e `frame` nel proprio blocco.

## <a id="constants-and-helpers"></a>Costanti e metodi di supporto

| Nome | Significato |
|---|---|
| `Docxi::VERSION` | La versione installata, per esempio `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Scrive una parte nell'archivio con la dimensione impostata in anticipo. Usato internamente da ogni parte |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Metodi di supporto interni con le regole di `blank?` di ActiveSupport |

## <a id="related-pages"></a>Pagine correlate

- [Funzionalità](Features)
- [Esempi d'uso](Usage-Examples)
- [Architettura](Architecture)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
