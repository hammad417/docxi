# Configuring docxi: page margins, page setup and units

docxi has no configuration file, no environment variables and no global settings. Everything is set per document or per element. This page covers the document-level settings and the units that options use. The options of each element are listed in [Features](Features).

## Page margins

Margins are passed when you create the document:

```ruby
require "docxi"

# All values are in twentieths of a point (1440 = 1 inch).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Option | Default | Meaning |
|---|---|---|
| `top` | `720` (0.5 in) | Top margin |
| `right` | `1440` (1 in) | Right margin |
| `bottom` | `720` (0.5 in) | Bottom margin |
| `left` | `1440` (1 in) | Left margin |
| `header` | `190` | Distance from the top edge of the page to the header |
| `footer` | `200` | Distance from the bottom edge of the page to the footer |

The hash you pass replaces the defaults, so set `top`, `right`, `bottom` and `left` together. Any of those four you leave out is written as an empty value. `header` and `footer` fall back to 190 and 200 when they're missing or blank.

## Page size and orientation

Pages are US Letter in portrait: 12240 × 15840 twentieths of a point, which is 8.5 × 11 inches. The size and orientation can't be changed in 0.1.0.

## First page

docxi always turns on Word's "different first page" setting. Headers and footers added with `add_header` and `add_footer` are used from page 2 onwards. Page 1 uses `add_first_header` and `add_first_footer`, so set those as well if the first page needs a header or footer. See [Features](Features#headers-and-footers).

## What every document contains

- The default font is Arial, 12 pt.
- Styles defined: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` and `NoList`.
- Bullet items (`ul: true`) use a dash bullet.
- The document properties name the application as "Docxi Ruby Gem" and record the time the document was rendered.

## Units

| Option | Unit |
|---|---|
| Text `size:` | Points (docxi doubles it into Word's half-points) |
| Margins, paragraph `left:` and `right:`, table `iwidth:`, row `height:` | Twentieths of a point (1440 per inch) |
| Cell `borders:` sizes | Eighths of a point |
| `columns_width:` and cell `width:` | Multiplied by 14.92 to get twentieths of a point, close to pixels at 96 dpi |
| Inline image `style:` | CSS lengths such as `"120pt"` |
| Positioned image `pH:` and `pV:` | EMU (914400 per inch) |
| Text and fill colors | Hex without `#`, for example `"8B0000"` |
| Image `fill_color:` | VML color with `#`, for example `"#00FF00"` |

## Related pages

- [Features](Features)
- [API Reference](API-Reference)
- [Best Practices](Best-Practices)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
