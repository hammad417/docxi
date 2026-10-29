<!-- Translated from wiki/Usage-Examples.md (git blob bf0201008119c1e93bee8c0df80b57836cc899d1). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Anwendungsbeispiele: Berichte, Rechnungen, Briefe und Tabellen in Ruby

Jedes Beispiel auf dieser Seite ist vollständig: Kopieren Sie es in eine Datei, führen Sie es mit Ruby aus und öffnen Sie die `.docx`-Datei, die es schreibt. Details zu den Optionen stehen unter [Funktionen](Features).

## <a id="a-report-with-a-cover-page-header-and-page-numbers"></a>Ein Bericht mit Kopfzeile auf dem Deckblatt und Seitenzahlen

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Internal report", size: 9, align: "right" })
document.add_header(Docxi::Word::Headers::Header.new { |h| h.text "Monthly sales", size: 9, align: "right" })
document.add_first_footer(Docxi::Word::Footers::Footer.new { |f| f.page_numbers name: "Page", align: "center" })
document.add_footer(Docxi::Word::Footers::Footer.new { |f| f.page_numbers name: "Page", align: "center" })

document.p(align: "center") { |p| p.text "Monthly sales report", size: 24, bold: true }
document.p(align: "center") { |p| p.text "September 2026", size: 14, color: "666666" }

document.br(page: true)

document.p { |p| p.text "Summary", size: 16, bold: true }
document.p do |p|
  p.text "Sales rose in every region. The table below lists the totals."
end

File.binwrite("report.docx", word.render.string)
```

Die erste Seite hat eine eigene Kopf- und Fußzeile, deshalb werden beide Paare gesetzt.

## <a id="an-invoice-with-a-totals-row"></a>Eine Rechnung mit Summenzeile

```ruby
require "docxi"

items = [
  { name: "Design work", qty: 10, price: 80.0 },
  { name: "Hosting (1 year)", qty: 1, price: 240.0 }
]

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Invoice #1042", size: 20, bold: true }
document.p { |p| p.text "Issued 2026-10-28", color: "666666" }

border = { top: 4, right: 4, bottom: 4, left: 4 }

document.table(columns_width: [260, 60, 110]) do |table|
  table.tr(fill: "D9D9D9", bold: true, borders: border) do |tr|
    tr.tc { |tc| tc.text "Item" }
    tr.tc { |tc| tc.text "Qty", align: "right" }
    tr.tc { |tc| tc.text "Amount", align: "right" }
  end

  items.each do |item|
    table.tr(borders: border) do |tr|
      tr.tc { |tc| tc.text item[:name] }
      tr.tc { |tc| tc.text item[:qty].to_s, align: "right" }
      tr.tc { |tc| tc.text format("$%.2f", item[:qty] * item[:price]), align: "right" }
    end
  end

  total = items.sum { |item| item[:qty] * item[:price] }
  table.tr(bold: true, borders: border) do |tr|
    tr.tc(colspan: 2) { |tc| tc.text "Total" }
    tr.tc { |tc| tc.text format("$%.2f", total), align: "right" }
  end
end

File.binwrite("invoice.docx", word.render.string)
```

Die Rahmen werden auf den Zeilen gesetzt, die sie an jede Zelle weitergeben.

## <a id="a-letter"></a>Ein Brief

```ruby
require "docxi"

word = Docxi::Document.new(top: 1440, right: 1440, bottom: 1440, left: 1440)
document = word.document

document.p(align: "right") do |p|
  p.text "ACME Corp.", bold: true, br: true
  p.text "12 Example Street", br: true
  p.text "2026-10-28"
end

document.p { |p| p.br }
document.p { |p| p.text "Dear Ms. Rivera," }
document.p(align: "both") do |p|
  p.text "Thank you for your order. Your documents are attached, and the next delivery is scheduled for November."
end
document.p { |p| p.text "Kind regards,", br: true }
document.p { |p| p.text "The ACME team", italic: true }

File.binwrite("letter.docx", word.render.string)
```

Alle vier Ränder werden zusammen übergeben. Das ist nötig, sobald Sie einen davon ändern. Siehe [Konfiguration](Configuration).

## <a id="a-table-built-from-records"></a>Eine Tabelle aus Datensätzen

```ruby
require "docxi"

people = [
  ["Ada Lovelace", "Engineering", "2019"],
  ["Grace Hopper", "Research", "2021"],
  ["Alan Turing", "Security", "2024"]
]

word = Docxi::Document.new
document = word.document

document.table(columns_width: [200, 150, 80]) do |table|
  table.tr(fill: "1F4E79", color: "FFFFFF", bold: true) do |tr|
    ["Name", "Team", "Since"].each { |heading| tr.tc { |tc| tc.text heading } }
  end

  people.each_with_index do |row, index|
    table.tr(fill: index.even? ? "FFFFFF" : "EEF3F8") do |tr|
      row.each { |value| tr.tc { |tc| tc.text value } }
    end
  end
end

File.binwrite("people.docx", word.render.string)
```

## <a id="a-document-with-an-image"></a>Ein Dokument mit einem Bild

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

chart = document.add_media("logo.png")
# Eine URL funktioniert ebenfalls:
# chart = document.add_media("https://example.com/chart.png")

document.p(align: "center") { |p| p.image chart, style: { width: "200pt", height: "100pt" } }
document.p(align: "center") { |p| p.text "Figure 1: sales by region", italic: true, size: 9 }

File.binwrite("figure.docx", word.render.string)
```

Führen Sie es neben einer `logo.png` aus. Wenn Sie einen Pfad statt eines geöffneten `File` übergeben, kann das Dokument mehrmals gerendert werden. Siehe [Best Practices](Best-Practices#images).

## <a id="a-multi-section-document-with-a-contents-page"></a>Ein Dokument mit mehreren Abschnitten und Inhaltsverzeichnis

```ruby
require "docxi"

sections = ["Introduction", "Results", "Next steps"]

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Contents", size: 16, bold: true }
document.table_of_content do |toc|
  sections.each_with_index { |title, index| toc.add_row title, index + 2 }
end

sections.each do |title|
  document.br(page: true)
  document.p { |p| p.text title, size: 16, bold: true }
  document.p { |p| p.text "Text of the #{title.downcase} section." }
end

File.binwrite("sections.docx", word.render.string)
```

Die Seitenzahlen im Inhaltsverzeichnis sind die, die Sie übergeben: Jeder Abschnitt beginnt auf einer neuen Seite, also steht Abschnitt `n` auf Seite `n + 1`.

## <a id="related-pages"></a>Verwandte Seiten

- [Erste Schritte](Getting-Started)
- [Integrationen](Integrations) für Rails, Rack und Hintergrundjobs
- [Fehlerbehebung](Troubleshooting)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
