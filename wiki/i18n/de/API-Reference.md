<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-API-Referenz: Klassen und Methoden

docxi ist eine Ruby-Bibliothek, daher listet diese Seite die Ruby-API auf. Eine HTTP-API gibt es nicht. Die Optionen jeder Methode sind unter [Funktionen](Features) beschrieben, die Klassen unter [Architektur](Architecture) erklärt.

## <a id="docxidocument"></a>Docxi::Document

| Methode | Rückgabe | Hinweise |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` sind die Seitenränder. Siehe [Konfiguration](Configuration) |
| `#document` | `Docxi::Word::Document` | Der Hauptteil, dem Sie Inhalte hinzufügen |
| `#render` | `StringIO` | Das vollständige `.docx`-Paket. `.string` liefert die Bytes |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Methode | Rückgabe | Fügt hinzu |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Einen Absatz |
| `text(string, options = {})` | Paragraph | Einen Absatz mit einem einzigen Textlauf |
| `br(page: true)` | Break | Einen Seitenumbruch |
| `tab(page: true)` | Tab | Einen Absatz mit einem Tabulator |
| `table(options = {}) { \|table\| }` | Table | Eine Tabelle |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Einträge des Inhaltsverzeichnisses |
| `frame(text, options = {})` | Paragraph | Einen Absatz mit einem Textfeld |
| `add_media(io_path_or_url, options = {})` | Media | Ein Bild, das mit `image` platziert wird |
| `add_hyperlink(url, options = {})` | String | Eine Beziehungs-ID für `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Kopfzeile ab Seite 2 und für Seite 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Fußzeile ab Seite 2 und für Seite 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Absatz (`Docxi::Word::Contents::Paragraph`)

| Methode | Fügt hinzu |
|---|---|
| `text(string, options = {})` | Einen Textlauf. Die Optionen des Absatzes sind seine Standardwerte |
| `br(times: 1)` | Zeilenumbrüche |
| `tab(times: 1)` | Tabulatorzeichen |
| `image(media, options = {})` | Ein Inline- oder positioniertes Bild |
| `hyperlink(text, url, id)` | Einen Link, der die Beziehungs-`id` aus `add_hyperlink` verwendet |
| `frame(options = {}) { \|frame\| }` | Ein Textfeld |

## <a id="tables"></a>Tabellen

| Methode | Fügt hinzu |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Eine Zeile. Die Tabellenoptionen sind ihre Standardwerte |
| `TableRow#tc(options = {}) { \|tc\| }` | Eine Zelle. Die Zeilenoptionen sind ihre Standardwerte |
| `TableCell#text(string, options = {})` | Einen Absatz. Die Zellenoptionen sind seine Standardwerte |
| `TableCell#p(options = {}) { \|p\| }` | Einen Absatz mit eigenen Optionen |
| `TableCell#image(media, width:, height:)` | Ein Bild |
| `TableCell#page_numbers` | Ein Seitenzahlfeld |

## <a id="table-of-contents"></a>Inhaltsverzeichnis

| Methode | Fügt hinzu |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Einen Eintrag im Inhaltsverzeichnis |

## <a id="headers-and-footers"></a>Kopf- und Fußzeilen

| Klasse | Methoden |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Rahmen

`Docxi::Word::Contents::Frame` akzeptiert in seinem Block `text`, `br`, `tab`, `image` und `frame`.

## <a id="constants-and-helpers"></a>Konstanten und Hilfsmethoden

| Name | Bedeutung |
|---|---|
| `Docxi::VERSION` | Die installierte Version, zum Beispiel `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Schreibt einen Teil ins Archiv und setzt seine Größe vorab. Wird intern von jedem Teil verwendet |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Interne Hilfsmethoden mit den `blank?`-Regeln von ActiveSupport |

## <a id="related-pages"></a>Verwandte Seiten

- [Funktionen](Features)
- [Anwendungsbeispiele](Usage-Examples)
- [Architektur](Architecture)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
