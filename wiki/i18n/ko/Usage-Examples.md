<!-- Translated from wiki/Usage-Examples.md (git blob bf0201008119c1e93bee8c0df80b57836cc899d1). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 사용 예제: Ruby로 보고서, 송장, 편지, 표 만들기

이 페이지의 예제는 모두 그대로 실행됩니다. 파일에 복사해 Ruby로 실행한 다음, 생성된 `.docx`를 열어 보세요. 옵션의 자세한 설명은 [기능](Features)에 있습니다.

## <a id="a-report-with-a-cover-page-header-and-page-numbers"></a>표지 머리글과 페이지 번호가 있는 보고서

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

첫 페이지에는 별도의 머리글과 바닥글이 있으므로 두 쌍을 모두 설정합니다.

## <a id="an-invoice-with-a-totals-row"></a>합계 행이 있는 송장

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

테두리는 행에 설정하고, 행이 이를 각 셀에 전달합니다.

## <a id="a-letter"></a>편지

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

네 여백을 함께 전달합니다. 여백을 하나라도 바꾸려면 이렇게 해야 합니다. [구성](Configuration)을 참고하세요.

## <a id="a-table-built-from-records"></a>레코드로 만드는 표

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

## <a id="a-document-with-an-image"></a>이미지가 있는 문서

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

chart = document.add_media("logo.png")
# URL도 사용할 수 있습니다:
# chart = document.add_media("https://example.com/chart.png")

document.p(align: "center") { |p| p.image chart, style: { width: "200pt", height: "100pt" } }
document.p(align: "center") { |p| p.text "Figure 1: sales by region", italic: true, size: 9 }

File.binwrite("figure.docx", word.render.string)
```

`logo.png`가 있는 위치에서 실행하세요. 열린 `File` 대신 경로를 전달하면 문서를 여러 번 렌더링할 수 있습니다. [모범 사례](Best-Practices#images)를 참고하세요.

## <a id="a-multi-section-document-with-a-contents-page"></a>목차 페이지가 있는 여러 섹션 문서

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

목차의 페이지 번호는 전달한 값이 그대로 쓰입니다. 각 섹션은 새 페이지에서 시작하므로, 섹션 `n`은 `n + 1`페이지에 있습니다.

## <a id="related-pages"></a>관련 페이지

- [시작하기](Getting-Started)
- [통합](Integrations): Rails, Rack, 백그라운드 작업
- [문제 해결](Troubleshooting)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
