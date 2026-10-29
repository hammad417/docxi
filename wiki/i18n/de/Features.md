<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Funktionen: Absätze, Text, Tabellen, Bilder, Kopf- und Fußzeilen

Diese Seite listet jedes Element auf, das docxi einem Word-Dokument hinzufügen kann, und die Optionen, die es jeweils akzeptiert. Die Beispiele bauen auf dieser Grundlage auf:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Optionen, die Sie nicht brauchen, können Sie weglassen. Farben sind Hex-Werte ohne `#`, zum Beispiel `"8B0000"`. Einheiten und Seiteneinrichtung stehen unter [Konfiguration](Configuration).

## <a id="paragraphs"></a>Absätze

`document.p(options) { |p| ... }` fügt einen Absatz hinzu.

| Option | Werte | Wirkung |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Ausrichtung. `"both"` setzt den Text im Blocksatz |
| `fill` | Hex-Farbe | Schattierung hinter dem Absatz und seinem Text |
| `left`, `right` | Zwanzigstelpunkt | Einzug. Setzen Sie beide, denn ein einzelner Wert wird ignoriert |
| `ul` | `true` | Macht den Absatz zu einem Aufzählungspunkt |
| `bottom` | `true` | Zeichnet eine Rahmenlinie unter dem Absatz |

Textoptionen wie `size:`, `bold:` oder `color:` können ebenfalls an `p` übergeben werden. Sie werden zu Standardwerten für jeden Textlauf im Absatz.

Innerhalb des Blocks können Sie `text`, `br`, `tab`, `image`, `hyperlink` und `frame` aufrufen. `document.text(string, options)` ist eine Abkürzung für einen Absatz mit einem einzigen Textlauf.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Text

`p.text(string, options)` fügt einen Textlauf (Run) hinzu.

| Option | Werte | Wirkung |
|---|---|---|
| `bold` | `true` | Fett |
| `italic` | `true` | Kursiv |
| `underline` | `true` | Einfach unterstrichen |
| `color` | Hex-Farbe | Textfarbe |
| `size` | Punkt | Schriftgröße |
| `font` | Schriftname | Schriftart, zum Beispiel `"Georgia"` |
| `fill` | Hex-Farbe | Schattierung hinter dem Text |
| `space` | `true` | Erhält Leerzeichen am Anfang und Ende des Texts |
| `br` | `true` | Fügt nach dem Text einen Zeilenumbruch ein |

Ein `nil`-String wird übersprungen.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Zeilenumbrüche, Seitenumbrüche und Tabulatoren

- `p.br` oder `p.br(times: 3)` fügt Zeilenumbrüche innerhalb eines Absatzes ein.
- `document.br(page: true)` beginnt eine neue Seite.
- `p.tab` oder `p.tab(times: 2)` fügt Tabulatorzeichen innerhalb eines Absatzes ein.
- `document.tab(page: true)` fügt einen Absatz mit einem einzelnen Tabulator ein.

Rufen Sie `document.br` oder `document.tab` nicht ohne `page: true` auf. Sie setzen dann einen Textlauf direkt in den Dokumentkörper, und das ist kein gültiges WordprocessingML.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tabellen

`document.table(options)` nimmt einen Block entgegen. Darin fügt `table.tr(options)` Zeilen hinzu und `tr.tc(options)` Zellen.

**Tabellenoptionen**

| Option | Werte | Wirkung |
|---|---|---|
| `columns_width` | Array von Zahlen | Spaltenbreiten, mit 14.92 in Zwanzigstelpunkt umgerechnet. Zugleich die Standardbreiten der Zellen |
| `height` | Zwanzigstelpunkt | Mindesthöhe jeder Zeile |
| `iwidth` | Zwanzigstelpunkt | Einzug der Tabelle vom linken Rand |
| `style` | Formatvorlagen-ID | Tabellenformatvorlage. docxi definiert `TableNormal` und `TableCustom` |
| `width` | Zahl | Bevorzugte Tabellenbreite, mit dem Typ `auto` geschrieben. Ohne `columns_width` erhält auch jede Zelle diese Breite, mit 14.92 multipliziert |
| `align` | `"left"`, `"center"`, `"right"` | Ausrichtung der Tabelle. Mit `tc.text` hinzugefügter Text wird genauso ausgerichtet |
| `padding` | `true` | Fügt feste Innenabstände in den Zellen hinzu |
| `borders` | Hash | Bekanntes Problem: wird auf jede Zelle angewendet statt auf den äußeren Rand der Tabelle |

**Zeilenoptionen**: `height:` legt die Mindesthöhe der Zeile fest. Jede Zellenoption, die an `tr` übergeben wird, gilt für alle Zellen der Zeile. Tabellenoptionen werden auf dieselbe Weise an jede Zeile und jede Zelle weitergegeben.

**Zellenoptionen**

| Option | Werte | Wirkung |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Rahmenstärke in Achtelpunkt. `nil` bedeutet keinen Rahmen auf dieser Seite |
| `fill` | Hex-Farbe | Zellschattierung. Zellen ohne Füllung sind weiß (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Vertikale Ausrichtung |
| `colspan` | Ganzzahl | Anzahl der Spalten, die die Zelle überspannt |
| `rowspan` | `true` | Beginnt eine vertikale Verbindung |
| `merged` | `true` | Setzt die vertikale Verbindung der Zelle darüber fort. Ihr Inhalt wird ignoriert |
| `width` | Zahl | Zellenbreite, mit 14.92 multipliziert |

**Zelleninhalt**

- `tc.text(string, options)`: Zellen- und Zeilenoptionen wie `bold:` oder `align:` gelten auch für den Text.
- `tc.p(options) { |p| ... }`: ein Absatz, der nur seine eigenen Optionen verwendet.
- `tc.image(media, width:, height:)`: ein Bild. Zum Größenproblem siehe unter [Bilder](#images).
- `tc.page_numbers`: ein Seitenzahlfeld. Bekanntes Problem: Ihm wird der Text „GlobalOptions“ vorangestellt.

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

## <a id="images"></a>Bilder

`media = document.add_media(source)` registriert ein Bild. `source` kann Folgendes sein:

- ein IO-Objekt, etwa ein `File` oder ein `StringIO`
- ein Dateipfad
- eine `http`-, `https`- oder `ftp`-URL

Das Bild wird beim Aufruf von `render` gelesen und immer als PNG-Teil gespeichert. PNG-Dateien sind daher die sichere Wahl.

**Inline-Bilder**: `p.image(media, options)` platziert das Bild in der Textzeile.

| Option | Werte | Wirkung |
|---|---|---|
| `style` | Hash mit CSS-Eigenschaften | Größe, zum Beispiel `{ width: "120pt", height: "60pt" }` |
| `title` | String | Titel des Bildes |
| `fill_color` | VML-Farbe wie `"#00FF00"` | Füllung hinter dem Bild |
| `opacity` | String wie `"50%"` | Deckkraft der Füllung |
| `wrap` | `true` | Schaltet den Textumbruch um das Bild aus |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Positionierte Bilder**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` verankert das Bild mit einem Versatz zur Spalte (`pH`) und zum Absatz (`pV`), in EMU. `pH`, `pV` und `wrap` müssen alle angegeben werden, sonst wird das Bild inline platziert.

**Bilder in Kopfzeilen**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` schaltet den Umbruch aus. Jeder andere Wert lässt den Text quadratisch um das Bild fließen.

**Bekanntes Problem in 0.1.0**: Positionierte Bilder, Bilder in Kopfzeilen und Bilder in Tabellenzellen berechnen Breite und Höhe beide als `width * height * 14.92` EMU. Sie sind immer quadratisch und meist winzig. Inline-Bilder, deren Größe über `style:` gesetzt wird, sind nicht betroffen. Einen Workaround finden Sie unter [Best Practices](Best-Practices#images).

## <a id="headers-and-footers"></a>Kopf- und Fußzeilen

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

- Eine Kopfzeile kann `text`, `br`, `table`, `image` und `page_numbers` enthalten. Eine Fußzeile kann `text`, `br`, `table` und `page_numbers` enthalten.
- Optionen, die an `Header.new` oder `Footer.new` übergeben werden, gelten für deren Text.
- `add_header` und `add_footer` gelten ab Seite 2. Seite 1 verwendet immer `add_first_header` und `add_first_footer`. Siehe [Konfiguration](Configuration#first-page).

## <a id="page-numbers"></a>Seitenzahlen

`page_numbers(name: "Page", align: "center")` schreibt den Text aus `name`, gefolgt von einem PAGE-Feld. Die Zahl erscheint in Arial 10 pt, dunkelgrau. `align` ist standardmäßig `"right"`.

Bekanntes Problem: In der Feldfunktion fehlt der Backslash von `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Inhaltsverzeichnis

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Mit Seitenzahl erhält der Eintrag einen rechtsbündigen Tabstopp bei `pos` (in Zwanzigstelpunkt; 9360 sind 6,5 Zoll), aufgefüllt mit dem Zeichen aus `leader`. Die Füllzeichenwerte stammen aus WordprocessingML, etwa `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` und `"none"`.
- Ohne Seitenzahl besteht der Eintrag aus dem Text nach einem Tabulator.
- Die Seitenzahlen sind die, die Sie übergeben. docxi erzeugt kein Word-TOC-Feld und berechnet keine Seiten.

## <a id="hyperlinks"></a>Hyperlinks

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` gibt eine Beziehungs-ID zurück, und `hyperlink(text, url, id)` setzt den Link. Das Ziel stammt aus der Beziehung, übergeben Sie also die ID, die für die gewünschte URL zurückgegeben wurde. Nur externe Links werden unterstützt.

## <a id="frames-text-boxes"></a>Rahmen (Textfelder)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Rahmen sind VML-Textfelder mit einer festen hellgrauen, halbtransparenten Füllung und einer schwarzen Kontur. Text fließt über und unter ihnen.
- In einem Rahmen können Sie `text`, `br`, `tab`, `image` und `frame` aufrufen.
- `spacing: true` fügt vor dem Text Abstand ein.

## <a id="output"></a>Ausgabe

`word.render` gibt ein `StringIO` mit dem vollständigen `.docx`-Paket zurück. Das Archiv wird ohne Zip64-Einträge geschrieben, sodass Google Docs und ältere LibreOffice-Versionen es öffnen können. [Architektur](Architecture) listet jeden Teil des Pakets auf.

## <a id="related-pages"></a>Verwandte Seiten

- [Anwendungsbeispiele](Usage-Examples)
- [API-Referenz](API-Reference)
- [Roadmap](Roadmap) mit den für 0.2.0 geplanten bekannten Problemen

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
