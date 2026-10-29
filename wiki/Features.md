# docxi features: paragraphs, text, tables, images, headers and footers

This page lists every element docxi can add to a Word document and the options each one accepts. The examples build on this setup:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Options you don't need can be left out. Colors are hex values without `#`, for example `"8B0000"`. For units and page setup, see [Configuration](Configuration).

## Paragraphs

`document.p(options) { |p| ... }` adds a paragraph.

| Option | Values | Effect |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Alignment. `"both"` justifies the text |
| `fill` | Hex color | Shading behind the paragraph and its text |
| `left`, `right` | Twentieths of a point | Indentation. Set both, because one on its own is ignored |
| `ul` | `true` | Makes the paragraph a bullet list item |
| `bottom` | `true` | Draws a border under the paragraph |

Text options such as `size:`, `bold:` or `color:` can also be given to `p`. They become the defaults for every run of text in the paragraph.

Inside the block you can call `text`, `br`, `tab`, `image`, `hyperlink` and `frame`. `document.text(string, options)` is a shortcut for a paragraph that holds a single run.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## Text

`p.text(string, options)` adds a run of text.

| Option | Values | Effect |
|---|---|---|
| `bold` | `true` | Bold |
| `italic` | `true` | Italic |
| `underline` | `true` | Single underline |
| `color` | Hex color | Text color |
| `size` | Points | Font size |
| `font` | Font name | Font, for example `"Georgia"` |
| `fill` | Hex color | Shading behind the text |
| `space` | `true` | Keeps spaces at the start and end of the text |
| `br` | `true` | Adds a line break after the text |

A `nil` string is skipped.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## Line breaks, page breaks and tabs

- `p.br`, or `p.br(times: 3)`, adds line breaks inside a paragraph.
- `document.br(page: true)` starts a new page.
- `p.tab`, or `p.tab(times: 2)`, adds tab characters inside a paragraph.
- `document.tab(page: true)` adds a paragraph that holds a single tab.

Don't call `document.br` or `document.tab` without `page: true`. They place a text run directly in the document body, which isn't valid WordprocessingML.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## Tables

`document.table(options)` takes a block, and inside it `table.tr(options)` adds rows and `tr.tc(options)` adds cells.

**Table options**

| Option | Values | Effect |
|---|---|---|
| `columns_width` | Array of numbers | Column widths, multiplied by 14.92 into twentieths of a point. Also the default cell widths |
| `height` | Twentieths of a point | Minimum height of every row |
| `iwidth` | Twentieths of a point | Indent of the table from the left margin |
| `style` | Style id | Table style. docxi defines `TableNormal` and `TableCustom` |
| `width` | Number | Preferred table width, written with the type `auto`. Without `columns_width`, every cell also gets this width, multiplied by 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Table alignment. Text added with `tc.text` is aligned the same way |
| `padding` | `true` | Adds fixed padding inside the cells |
| `borders` | Hash | Known issue: applied to every cell rather than to the outside of the table |

**Row options**: `height:` sets the minimum height of the row. Any cell option given to `tr` applies to every cell in the row. Table options are passed down the same way, to every row and cell.

**Cell options**

| Option | Values | Effect |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Border width in eighths of a point. `nil` means no border on that side |
| `fill` | Hex color | Cell shading. Cells without a fill are white (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Vertical alignment |
| `colspan` | Integer | Number of columns the cell spans |
| `rowspan` | `true` | Starts a vertical merge |
| `merged` | `true` | Continues a vertical merge from the cell above. Its content is ignored |
| `width` | Number | Cell width, multiplied by 14.92 |

**Cell content**

- `tc.text(string, options)`: cell and row options, such as `bold:` or `align:`, also apply to the text.
- `tc.p(options) { |p| ... }`: a paragraph with its own options only.
- `tc.image(media, width:, height:)`: an image. See the sizing issue under [Images](#images).
- `tc.page_numbers`: a page number field. Known issue: it's prefixed with the text "GlobalOptions".

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

## Images

`media = document.add_media(source)` registers an image. `source` can be:

- an IO object, such as a `File` or a `StringIO`
- a file path
- an `http`, `https` or `ftp` URL

The image is read when you call `render`, and is always stored as a PNG part, so PNG files are the safe choice.

**Inline images**: `p.image(media, options)` places the image in the line of text.

| Option | Values | Effect |
|---|---|---|
| `style` | Hash of CSS properties | Size, for example `{ width: "120pt", height: "60pt" }` |
| `title` | String | Title of the image |
| `fill_color` | VML color such as `"#00FF00"` | Fill behind the image |
| `opacity` | String such as `"50%"` | Opacity of the fill |
| `wrap` | `true` | Turns off text wrapping around the image |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Positioned images**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` anchors the image at an offset from the column (`pH`) and the paragraph (`pV`), in EMU. `pH`, `pV` and `wrap` must all be given, otherwise the image is placed inline.

**Header images**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` turns wrapping off. Any other value wraps text around a square.

**Known issue in 0.1.0**: positioned images, header images and table-cell images compute both their width and their height as `width * height * 14.92` EMU. They always come out square and usually tiny. Inline images sized with `style:` are not affected. See [Best Practices](Best-Practices#images) for a workaround.

## Headers and footers

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

- A header can hold `text`, `br`, `table`, `image` and `page_numbers`. A footer can hold `text`, `br`, `table` and `page_numbers`.
- Options passed to `Header.new` or `Footer.new` apply to their text.
- `add_header` and `add_footer` cover pages 2 onwards. Page 1 always uses `add_first_header` and `add_first_footer`. See [Configuration](Configuration#first-page).

## Page numbers

`page_numbers(name: "Page", align: "center")` writes the `name` text followed by a PAGE field. The number is set in Arial 10 pt, dark grey. `align` defaults to `"right"`.

Known issue: the field code is missing the backslash in `\* MERGEFORMAT`.

## Table of contents

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- With a page number, the row gets a right-aligned tab stop at `pos` (twentieths of a point; 9360 is 6.5 inches) filled with the `leader` character. Leader values come from WordprocessingML, such as `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` and `"none"`.
- Without a page number, the row is the text after a tab.
- The page numbers are the ones you pass. docxi doesn't create a Word TOC field and doesn't calculate pages.

## Hyperlinks

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` returns a relationship id, and `hyperlink(text, url, id)` places the link. The target comes from the relationship, so pass the id returned for the URL you want. Only external links are supported.

## Frames (text boxes)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Frames are VML text boxes with a fixed light grey, semi-transparent fill and a black outline. Text wraps above and below them.
- Inside a frame you can call `text`, `br`, `tab`, `image` and `frame`.
- `spacing: true` adds space before the text.

## Output

`word.render` returns a `StringIO` with the complete `.docx` package. The archive is written without Zip64 records, so Google Docs and older LibreOffice versions can open it. [Architecture](Architecture) lists every part of the package.

## Related pages

- [Usage Examples](Usage-Examples)
- [API Reference](API-Reference)
- [Roadmap](Roadmap) for the known issues planned for 0.2.0

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
