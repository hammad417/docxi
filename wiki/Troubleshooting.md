# Troubleshooting docxi: installation errors, runtime errors and broken output

Find your symptom below. Most problems with older apps come from docxi 0.0.x, and upgrading to 0.1.0 fixes them without code changes.

## Installation

**Bundler can't find compatible versions for `nokogiri` in a Rails 6+ app.**
docxi 0.0.6 pins `nokogiri < 1.7`, and every Rails since 6.0 needs nokogiri 1.8.5 or newer. Use `gem "docxi", "~> 0.1"`.

**Bundler can't find compatible versions for `rubyzip`.**
docxi 0.0.6 pins `rubyzip ~> 0.9.9`, which conflicts with any gem that needs a newer rubyzip. docxi 0.1.0 accepts rubyzip 1.3 to 3.x.

**Bundler keeps installing 0.0.6.**
Either the Gemfile pins docxi to 0.0.x, for example with `"~> 0.0.6"`, or Ruby is older than 2.5, which 0.1.0 requires. Change the Gemfile line to `gem "docxi", "~> 0.1"` and check `ruby -v`.

**nokogiri fails to compile.**
No precompiled nokogiri gem matches your platform, so it's building from source. Install a C compiler and follow the [nokogiri installation guide](https://nokogiri.org/tutorials/installing_nokogiri.html).

## Errors while building a document

**`NoMethodError: undefined method 'present?'` outside Rails.**
docxi 0.0.6 assumed ActiveSupport was loaded. 0.1.0 doesn't need it.

**`NameError` when calling `tab`, `frame` or a header's `page_numbers`.**
These were broken in 0.0.6 and work in 0.1.0, including `document.frame("text")`.

**`Errno::ENOENT` when adding an image.**
The path doesn't exist, relative to the directory your program runs in. On Ruby 3 with docxi 0.0.6, image URLs also fail this way, because 0.0.6 used `Kernel#open`. 0.1.0 downloads URLs with `URI#open`.

## Documents that don't look right

**Google Docs or LibreOffice won't open the file.**
docxi 0.0.6 with rubyzip 3 writes Zip64 records into every entry, and these programs reject them. 0.1.0 writes archives without them.

**The first page has no header or footer.**
docxi gives the first page its own header and footer. Set them with `add_first_header` and `add_first_footer`. See [Configuration](Configuration#first-page).

**A margin is missing, or the page layout looks wrong.**
When you set margins, pass `top`, `right`, `bottom` and `left` together. The ones you leave out are written empty.

**Spaces between runs of text disappear.**
Word drops spaces at the start or end of a run. Add `space: true` to that `text` call.

**Images in headers, table cells or with `pH`/`pV` are tiny or square.**
Known issue in 0.1.0: their size is computed as `width * height * 14.92` EMU for both dimensions. Use inline images sized with `style:`, or see the workaround in [Best Practices](Best-Practices#images).

**An image is missing the second time the same document is rendered.**
Images added from an open `File` or other IO are read during `render`, and a second render reads an empty stream. Pass a path instead, or render once.

**Borders set on `table` appear around every cell.**
Known issue: table-level `borders:` are applied to every cell rather than to the outside of the table. To draw only the outline, set `borders:` on the cells along the edges, for their outer sides only.

**Page numbers in table cells start with "GlobalOptions".**
Known issue in `TableCell#page_numbers`. Put page numbers in a footer or header instead.

**A JPEG or GIF image looks wrong or doesn't appear.**
Every image is stored as a PNG part with a PNG content type, whatever its real format, so some programs may not display other formats. Convert images to PNG before adding them.

## Inspecting a generated file

A `.docx` is a zip archive, so you can look inside it:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

If the problem is still there, [open an issue](https://github.com/hammad417/docxi/issues) with a short script that reproduces it and the versions of Ruby, docxi, nokogiri and rubyzip.

## Related pages

- [FAQ](FAQ)
- [Best Practices](Best-Practices)
- [Installation](Installation)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
