<!-- Translated from wiki/Usage-Examples.md (git blob bf0201008119c1e93bee8c0df80b57836cc899d1). See wiki/TRANSLATION_GUIDE.md. -->
# Примеры использования docxi: отчёты, счета, письма и таблицы на Ruby

Каждый пример на этой странице завершён: скопируйте его в файл, запустите Ruby и откройте созданный `.docx`. Подробности о параметрах — в разделе [Возможности](Features).

## <a id="a-report-with-a-cover-page-header-and-page-numbers"></a>Отчёт с колонтитулом титульной страницы и номерами страниц

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

У первой страницы свои колонтитулы, поэтому задаются обе пары.

## <a id="an-invoice-with-a-totals-row"></a>Счёт со строкой итогов

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

Границы задаются для строк, а строки передают их каждой ячейке.

## <a id="a-letter"></a>Письмо

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

Все четыре поля передаются вместе, и это обязательно, если вы меняете любое из них. См. [Настройка](Configuration).

## <a id="a-table-built-from-records"></a>Таблица из записей

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

## <a id="a-document-with-an-image"></a>Документ с изображением

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

chart = document.add_media("logo.png")
# URL тоже подойдёт:
# chart = document.add_media("https://example.com/chart.png")

document.p(align: "center") { |p| p.image chart, style: { width: "200pt", height: "100pt" } }
document.p(align: "center") { |p| p.text "Figure 1: sales by region", italic: true, size: 9 }

File.binwrite("figure.docx", word.render.string)
```

Запускайте пример рядом с файлом `logo.png`. Если передать путь, а не открытый `File`, документ можно будет формировать несколько раз. См. [Рекомендации](Best-Practices#images).

## <a id="a-multi-section-document-with-a-contents-page"></a>Документ из нескольких разделов со страницей оглавления

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

Номера страниц в оглавлении — те, что вы передаёте: каждый раздел начинается с новой страницы, поэтому раздел `n` находится на странице `n + 1`.

## <a id="related-pages"></a>Связанные страницы

- [Начало работы](Getting-Started)
- [Интеграции](Integrations): Rails, Rack и фоновые задачи
- [Устранение неполадок](Troubleshooting)

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
