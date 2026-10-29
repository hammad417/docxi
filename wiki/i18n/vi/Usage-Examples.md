<!-- Translated from wiki/Usage-Examples.md (git blob bf0201008119c1e93bee8c0df80b57836cc899d1). See wiki/TRANSLATION_GUIDE.md. -->
# Ví dụ sử dụng docxi: báo cáo, hóa đơn, thư và bảng bằng Ruby

Mỗi ví dụ trên trang này đều hoàn chỉnh: sao chép vào một tệp, chạy bằng Ruby rồi mở tệp `.docx` mà ví dụ ghi ra. Chi tiết về các tùy chọn có trong [Tính năng](Features).

## <a id="a-report-with-a-cover-page-header-and-page-numbers"></a>Báo cáo có đầu trang cho trang bìa và số trang

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

Trang đầu có đầu trang và chân trang riêng, vì vậy cả hai cặp đều được đặt.

## <a id="an-invoice-with-a-totals-row"></a>Hóa đơn có hàng tổng cộng

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

Viền được đặt trên các hàng, và hàng chuyển viền xuống từng ô.

## <a id="a-letter"></a>Một lá thư

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

Cả bốn lề được truyền cùng lúc, và đây là điều bắt buộc khi bạn thay đổi bất kỳ lề nào. Xem [Cấu hình](Configuration).

## <a id="a-table-built-from-records"></a>Bảng dựng từ các bản ghi

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

## <a id="a-document-with-an-image"></a>Tài liệu có hình ảnh

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

chart = document.add_media("logo.png")
# Cũng có thể dùng URL:
# chart = document.add_media("https://example.com/chart.png")

document.p(align: "center") { |p| p.image chart, style: { width: "200pt", height: "100pt" } }
document.p(align: "center") { |p| p.text "Figure 1: sales by region", italic: true, size: 9 }

File.binwrite("figure.docx", word.render.string)
```

Hãy chạy ví dụ trong thư mục có tệp `logo.png`. Truyền đường dẫn thay vì một `File` đang mở giúp tài liệu có thể được kết xuất nhiều lần. Xem [Thực hành tốt](Best-Practices#images).

## <a id="a-multi-section-document-with-a-contents-page"></a>Tài liệu nhiều phần có trang mục lục

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

Số trang trong mục lục là những số bạn truyền vào: mỗi phần bắt đầu ở một trang mới, nên phần `n` nằm ở trang `n + 1`.

## <a id="related-pages"></a>Trang liên quan

- [Bắt đầu](Getting-Started)
- [Tích hợp](Integrations): Rails, Rack và tác vụ nền
- [Khắc phục sự cố](Troubleshooting)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
