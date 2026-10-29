<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 快速入门：用 Ruby 生成第一个 Word 文档

本指南带你从一个空文件夹开始，生成一个带格式的 `.docx` 文件。你需要了解 Ruby 基础，但不需要 Rails 或 Microsoft Word。

## <a id="1-install-the-gem"></a>1. 安装 gem

```sh
gem install docxi
```

如果项目使用 Gemfile，请改为添加 `gem "docxi", "~> 0.1"` 并运行 `bundle install`。详情参见[安装](Installation)。

## <a id="2-create-a-document"></a>2. 创建文档

将以下代码保存为 `hello.rb`：

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

运行 `ruby hello.rb`（在 Bundler 项目中运行 `bundle exec ruby hello.rb`），然后打开 `hello.docx`。

## <a id="3-what-each-line-does"></a>3. 每一行的作用

- `Docxi::Document.new` 创建一个使用默认页边距的空文档。
- `word.document` 是文档的正文，所有内容都通过它添加。
- `document.p` 添加一个段落。传给 `p` 的选项（例如 `align:`）同时作用于段落及其中的文本。
- `p.text` 添加一个文本片段（run）。`size:` 的单位是磅。
- `word.render` 构建 `.docx` 包并返回一个 `StringIO`，它的 `.string` 就是文件的字节内容。

## <a id="4-add-a-list-and-a-table"></a>4. 添加列表和表格

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

有两点需要注意：

- `space: true` 会保留文本片段开头或结尾的空格。如果不加，Word 会丢掉这个空格，前后的词就会连在一起。
- `tr` 上的选项作用于该行的每个单元格，所以整个表头行都带底纹并加粗。

## <a id="5-where-to-go-next"></a>5. 下一步

- [使用示例](Usage-Examples)：报告、发票和信函的完整示例
- [功能](Features)：段落、文本、表格、图片、页眉和页脚的所有选项
- [集成](Integrations)：从 Rails 控制器发送文档
- [最佳实践](Best-Practices)：提前避开常见的意外

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
