<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 快速入門：用 Ruby 產生第一份 Word 文件

本指南帶你從一個空資料夾開始，產生一個帶格式的 `.docx` 檔案。你需要具備 Ruby 基礎，但不需要 Rails 或 Microsoft Word。

## <a id="1-install-the-gem"></a>1. 安裝 gem

```sh
gem install docxi
```

如果專案使用 Gemfile，請改為加入 `gem "docxi", "~> 0.1"` 並執行 `bundle install`。詳情請參閱[安裝](Installation)。

## <a id="2-create-a-document"></a>2. 建立文件

將以下程式碼儲存為 `hello.rb`：

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

執行 `ruby hello.rb`（在 Bundler 專案中執行 `bundle exec ruby hello.rb`），然後開啟 `hello.docx`。

## <a id="3-what-each-line-does"></a>3. 每一行的作用

- `Docxi::Document.new` 建立一份使用預設邊界的空白文件。
- `word.document` 是文件的本文，所有內容都透過它加入。
- `document.p` 加入一個段落。傳給 `p` 的選項（例如 `align:`）同時作用於段落及其中的文字。
- `p.text` 加入一個文字片段（run）。`size:` 的單位是點。
- `word.render` 建立 `.docx` 封裝並傳回一個 `StringIO`，它的 `.string` 就是檔案的位元組內容。

## <a id="4-add-a-list-and-a-table"></a>4. 加入清單和表格

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

有兩點需要注意：

- `space: true` 會保留文字片段開頭或結尾的空格。如果不加，Word 會捨棄這個空格，前後的字就會黏在一起。
- `tr` 上的選項會套用到該列的每個儲存格，所以整個標題列都有網底並以粗體顯示。

## <a id="5-where-to-go-next"></a>5. 下一步

- [使用範例](Usage-Examples)：報告、發票和信件的完整範例
- [功能](Features)：段落、文字、表格、圖片、頁首和頁尾的所有選項
- [整合](Integrations)：從 Rails 控制器傳送文件
- [最佳做法](Best-Practices)：事先避開常見的意外

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
