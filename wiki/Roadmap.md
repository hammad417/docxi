# docxi roadmap and known issues

docxi has no release schedule or committed roadmap. The only planned work on record is fixing the known issues below. They're planned for 0.2.0, without a date. They're kept out of 0.1.0 because fixing them changes the documents existing users already generate.

## Known issues planned for 0.2.0

| Issue | Workaround in 0.1.0 |
|---|---|
| Positioned, header and table-cell images compute their size as `width * height * 14.92` EMU for both dimensions | Use inline images sized with `style:`, or see [Best Practices](Best-Practices#images) |
| Table-level `borders:` are applied to every cell rather than to the outside of the table | For an outline only, set `borders:` on the outer sides of the edge cells |
| The page-number field code is missing the backslash in `\* MERGEFORMAT` | None |
| Page numbers in table cells are prefixed with "GlobalOptions" | Put page numbers in a header or footer |
| Images are always stored as PNG parts, and `[Content_Types].xml` only declares the first two headers and footers | Use PNG images, and at most two headers and two footers |
| Passing only some margins writes the others empty | Pass `top`, `right`, `bottom` and `left` together |
| `docProps/core.xml` records local time with a UTC marker and never writes the author | None |
| `Docxi.test` raises `NoMethodError` | Don't call it. It's a leftover helper |
| `document.br` and `document.tab` without `page: true` produce invalid XML | Use `p.br` and `p.tab` inside a paragraph |

## Not planned

These aren't on the list above, and nothing has been decided about them:

- editing existing documents or filling in templates
- numbered lists, heading styles or an automatic table of contents
- other page sizes or landscape pages

## Related pages

- [Changelog](Changelog)
- [Contributing](Contributing)
- [FAQ](FAQ)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
