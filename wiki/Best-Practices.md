# docxi best practices: avoid the common surprises

These habits come straight from how docxi works, and most of them prevent a problem listed in [Troubleshooting](Troubleshooting). They're grouped by topic. The page setup advice applies to everyone, and the security and performance advice matters most for production apps.

## Page setup

- **Pass all four margins together.** The hash you pass to `Docxi::Document.new` replaces the defaults, and any of `top`, `right`, `bottom` and `left` you leave out is written empty.
- **Set the first-page header and footer.** docxi always turns on a separate first page. If page 1 needs a header or footer, call `add_first_header` and `add_first_footer` as well.

## Text

- **Use `space: true` for leading or trailing spaces.** Otherwise Word drops them and the words of two runs touch.
- **Put shared formatting on the paragraph.** Options given to `p`, such as `size:` or `font:`, apply to every run inside it, so there's less to repeat.
- **Use `p.br` for line breaks and `document.br(page: true)` for page breaks.** `document.br` without `page: true` produces XML that isn't valid.

## Tables

- **Set borders on rows or cells, not on the table.** In 0.1.0, table-level `borders:` are applied to every cell rather than to the outside of the table.
- **Give `columns_width` for every column.** Cells take their width from it, so tables keep the layout you expect.

## Images

- **Use PNG files.** Every image is stored as a PNG part.
- **Size inline images with `style:`**, for example `style: { width: "120pt", height: "60pt" }`. It's the one image size that works as you'd expect in 0.1.0.
- **Pass a path rather than an open `File` if the document might be rendered more than once.** IO images are read during `render`, and a second render gets an empty image.
- **Header, table-cell and positioned images until 0.2.0.** Their width and height are both computed as `width * height * 14.92` EMU, so they come out square. To get a square image of `S` inches, pass `width` and `height` equal to `Math.sqrt(S * 914400 / 14.92).round`, which is about 248 for one inch.

## Security

- **Don't pass untrusted strings to `add_media`.** A path makes docxi read that file from your server, and a URL makes your server download it. Pass an IO you opened yourself, or check the value first. See [Security](Security).
- **Text is escaped for you.** Strings passed to `text` are written as XML text, so user input can't change the document structure.

## Performance and reliability

- **Build big documents in a background job.** The whole document is built in memory when you call `render`. See [Integrations](Integrations).
- **Render once, then reuse the bytes.** Keep `word.render.string` if you need to send or store the document more than once.
- **Pin a compatible version** with `gem "docxi", "~> 0.1"`.
- **Test what you generate.** A `.docx` is a zip of XML files, so tests can unzip it and check the XML. docxi's own tests work this way. See [Architecture](Architecture).

## Related pages

- [Features](Features)
- [Configuration](Configuration)
- [Troubleshooting](Troubleshooting)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
