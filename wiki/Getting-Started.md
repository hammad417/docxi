# Getting started with docxi: your first Word document in Ruby

This guide takes you from an empty folder to a formatted `.docx` file. It assumes you know basic Ruby. You don't need Rails or Microsoft Word.

## 1. Install the gem

```sh
gem install docxi
```

In a project with a Gemfile, add `gem "docxi", "~> 0.1"` and run `bundle install` instead. [Installation](Installation) has the details.

## 2. Create a document

Save this as `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Run `ruby hello.rb` (or `bundle exec ruby hello.rb` in a Bundler project) and open `hello.docx`.

## 3. What each line does

- `Docxi::Document.new` creates an empty document with the default page margins.
- `word.document` is the body of the document. Everything you add goes through it.
- `document.p` adds a paragraph. Options given to `p`, such as `align:`, apply to the paragraph and to the text inside it.
- `p.text` adds a run of text. `size:` is in points.
- `word.render` builds the `.docx` package and returns a `StringIO`. Its `.string` holds the bytes of the file.

## 4. Add a list and a table

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

Two things to notice:

- `space: true` keeps the space at the start or end of a run. Without it, Word drops the space and the words run together.
- Options on `tr` apply to every cell in that row, so the whole header row is shaded and bold.

## 5. Where to go next

- [Usage Examples](Usage-Examples): complete recipes for reports, invoices and letters
- [Features](Features): every option for paragraphs, text, tables, images, headers and footers
- [Integrations](Integrations): sending the document from a Rails controller
- [Best Practices](Best-Practices): avoid the common surprises before they happen

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
