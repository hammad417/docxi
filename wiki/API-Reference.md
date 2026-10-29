# docxi API reference: classes and methods

docxi is a Ruby library, so this page lists its Ruby API. There is no HTTP API. Options for each method are described in [Features](Features), and the classes are explained in [Architecture](Architecture).

## Docxi::Document

| Method | Returns | Notes |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` are page margins. See [Configuration](Configuration) |
| `#document` | `Docxi::Word::Document` | The body you add content to |
| `#render` | `StringIO` | The complete `.docx` package. `.string` gives the bytes |

## Docxi::Word::Document (`word.document`)

| Method | Returns | Adds |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | A paragraph |
| `text(string, options = {})` | Paragraph | A paragraph with one run of text |
| `br(page: true)` | Break | A page break |
| `tab(page: true)` | Tab | A paragraph holding one tab |
| `table(options = {}) { \|table\| }` | Table | A table |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Contents rows |
| `frame(text, options = {})` | Paragraph | A paragraph holding a text box |
| `add_media(io_path_or_url, options = {})` | Media | An image to place with `image` |
| `add_hyperlink(url, options = {})` | String | A relationship id for `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Header for pages 2 onwards, and for page 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Footer for pages 2 onwards, and for page 1 |

## Paragraph (`Docxi::Word::Contents::Paragraph`)

| Method | Adds |
|---|---|
| `text(string, options = {})` | A run of text. The paragraph's options are its defaults |
| `br(times: 1)` | Line breaks |
| `tab(times: 1)` | Tab characters |
| `image(media, options = {})` | An inline or positioned image |
| `hyperlink(text, url, id)` | A link that uses the relationship `id` from `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | A text box |

## Tables

| Method | Adds |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | A row. Table options are its defaults |
| `TableRow#tc(options = {}) { \|tc\| }` | A cell. Row options are its defaults |
| `TableCell#text(string, options = {})` | A paragraph. Cell options are its defaults |
| `TableCell#p(options = {}) { \|p\| }` | A paragraph with its own options |
| `TableCell#image(media, width:, height:)` | An image |
| `TableCell#page_numbers` | A page number field |

## Table of contents

| Method | Adds |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | One contents row |

## Headers and footers

| Class | Methods |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## Frames

`Docxi::Word::Contents::Frame` accepts `text`, `br`, `tab`, `image` and `frame` inside its block.

## Constants and helpers

| Name | Meaning |
|---|---|
| `Docxi::VERSION` | The installed version, for example `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Writes one part into the archive with its size set up front. Used internally by every part |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Internal helpers with ActiveSupport's `blank?` rules |

## Related pages

- [Features](Features)
- [Usage Examples](Usage-Examples)
- [Architecture](Architecture)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
