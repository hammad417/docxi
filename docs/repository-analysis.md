# Code notes

How docxi is put together and how it behaves, for anyone changing the code or the docs.

docxi builds Microsoft Word `.docx` files from Ruby code. You describe the document with a block-based DSL, and `Docxi::Document#render` returns a `StringIO` holding the Office Open XML package. It doesn't read or edit existing files, fill in templates or run a mail merge.

It started in 2015 as docxer by Timur Gabdrakipov (MIT). It was renamed docxi in 2016 and published as 0.0.1, 0.0.2 and 0.0.6. The 0.1 releases from 2026 bring it up to date with current Ruby, Rails and rubyzip without changing the documents it makes.

## Layout

```text
lib/docxi.rb                      entry point, Docxi.to_xml, Docxi.write_entry
lib/docxi/document.rb             Docxi::Document: package parts and #render
lib/docxi/util.rb                 blank?/present? helpers (replace ActiveSupport)
lib/docxi/content_types.rb        [Content_Types].xml
lib/docxi/relationships.rb        _rels/.rels
lib/docxi/properties/             docProps/app.xml and core.xml
lib/docxi/word/document.rb        Docxi::Word::Document: body, sections, parts
lib/docxi/word/helpers.rb         p, text, br, tab, frame, table, table_of_content
lib/docxi/word/contents/          Paragraph, Text, Break, Tab, Frame, Image, Table, TableOfContent
lib/docxi/word/headers/, footers/ headers, footers and their page number fields
lib/docxi/word/medias/            images added with add_media
lib/docxi/word/hyperlinks/        hyperlink relationships
lib/docxi/word/*.rb               styles, numbering, settings, fonts, theme, effects, notes
test/                             Minitest suite, golden files, sample documents
gemfiles/                         Gemfiles for each Rails version and dependency floor in CI
```

Most of the larger files (`styles.rb`, `themes/theme.rb`, `effects.rb`) are fixed XML.

## Dependencies

Ruby `>= 2.5`, nokogiri `>= 1.8, < 2` for the XML, and rubyzip `>= 1.3, < 4` for the zip package. There is no Rails dependency, no configuration file or environment variable, and no network access unless `add_media` is given a URL.

## Public API

| Entry point | Defined in |
|---|---|
| `Docxi::Document.new(options)` and `#render` | `lib/docxi/document.rb` |
| `#document` returns `Docxi::Word::Document` | `lib/docxi/document.rb` |
| `p`, `text`, `br`, `tab`, `frame`, `table`, `table_of_content` | `lib/docxi/word/helpers.rb` |
| `add_media`, `add_hyperlink`, `add_header`, `add_first_header`, `add_footer`, `add_first_footer` | `lib/docxi/word/document.rb` |
| Paragraph `text`, `br`, `tab`, `image`, `hyperlink`, `frame` | `lib/docxi/word/contents/paragraph.rb` |
| Table `tr`; row `tc`; cell `text`, `p`, `image`, `page_numbers` | `lib/docxi/word/contents/table.rb` |
| `Docxi::Word::Headers::Header` (`text`, `br`, `table`, `image`, `page_numbers`) | `lib/docxi/word/headers/header.rb` |
| `Docxi::Word::Footers::Footer` (`text`, `br`, `table`, `page_numbers`) | `lib/docxi/word/footers/footer.rb` |

## Behaviour worth knowing

- Margins are in twentieths of a point. The defaults are `top: 720, right: 1440, bottom: 720, left: 1440, header: 190, footer: 200`, and the hash you pass replaces them, so margins you leave out are written empty.
- The page is US Letter portrait (12240 × 15840) and `w:titlePg` is always set, so page 1 uses the first-page header and footer.
- The default font is Arial 12 pt. Text `size` is in points and doubled into half-points. Borders are in eighths of a point. `columns_width` and cell `width` are multiplied by 14.92.
- Options cascade. `Table#tr` merges the table's options into each row's, and `TableRow#tc` merges the row's into each cell's. So table-level `borders`, `fill`, `width` and `align` also apply to every cell, and `align` to the text of each `tc.text` paragraph. Table `width` is written as `w:tblW` with `w:type="auto"`, and without `columns_width` each cell gets that width × 14.92.
- Every zip entry goes through `Docxi.write_entry`, which sets the size first. Without it, rubyzip 3 adds Zip64 records that Google Docs and older LibreOffice reject.

## Known issues

These are planned for 0.2.0. They're left alone in 0.1.x because fixing them changes documents people already generate.

- Positioned, header and table-cell images compute both dimensions as `width * height * 14.92` EMU.
- Table-level `borders:` are written into `w:tblBorders` as `w:val="none"`. Because of the cascade above, the borders show up around every cell instead of around the table.
- The page-number field code is `PAGE   * MERGEFORMAT`. Ruby drops the backslash from the `"\*"` escape.
- Table-cell page numbers print the hard-coded text "GlobalOptions".
- Images are always saved as `media/imageN.png` with a PNG content type, and `[Content_Types].xml` only declares header1–2 and footer1–2.
- `docProps/core.xml` stamps local time with a `Z` suffix and never writes the creator.
- `Docxi.test` raises `NoMethodError` because it calls `document.break`.
- Document-level `br` and `tab` without `page: true` put a run directly in the body, which isn't valid WordprocessingML.

Not implemented: editing existing documents, numbered lists, heading styles, a real TOC field, other page sizes, multiple sections, footnote content, comments and charts.

## Tests and CI

| Command | What it does |
|---|---|
| `bundle exec rake test` | Runs the Minitest suite (the default task) |
| `bundle exec rake golden` | Rewrites `test/golden` from the current code |
| `bundle exec rake samples` | Writes the sample documents to `tmp/samples/` |

`test/golden` holds every part that 0.0.6 generated for the two documents in `test/support/sample_documents.rb`, and the tests compare each part as canonical XML. Only regenerate it after an intentional output change.

`ci.yml` runs the suite on Ruby 2.5 to 4.0, with Rails 5.2 to 8.1 loaded, with rubyzip 1.3 and 2.4 and nokogiri 1.11, and converts the samples to PDF with LibreOffice. `release.yml` publishes to RubyGems through trusted publishing when a `v*` tag is pushed. `docs.yml` checks the documentation.
