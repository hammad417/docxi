# docxi FAQ: Word documents in Ruby and Rails

Short answers to the questions people ask most. Longer explanations are linked from each answer.

**Do I need Microsoft Word or LibreOffice installed?**
No. docxi writes the XML and the zip archive itself.

**Does docxi require Rails?**
No. It works in any Ruby program. The tests also run with Rails 5.2 through 8.1 loaded. See [Integrations](Integrations).

**How do I generate a Word document in Ruby on Rails?**
Build it in a controller or a background job and send `word.render.string` with `send_data`. See [Integrations](Integrations).

**Which Ruby and Rails versions are supported?**
Ruby 2.5 and newer. Rails is optional, and CI tests Rails 5.2 through 8.1. See [Installation](Installation).

**Can docxi edit an existing .docx or fill in a Word template?**
No, it only creates new documents. For templates with merge fields, use a template-based gem such as sablon.

**Is the table of contents generated from headings?**
No. You add each row and its page number. See [Features](Features#table-of-contents).

**Can I use numbered lists or heading styles?**
Not in 0.1.0. Paragraphs can be bullet items (`ul: true`), and you can format headings with text options such as `size:` and `bold:`.

**Can I change the page size or orientation?**
Not yet. Pages are US Letter in portrait, and you can only change the margins. See [Configuration](Configuration).

**Which image formats work?**
PNG. Every image is stored as a PNG part.

**Can I add an image from a URL?**
Yes. `add_media` accepts `http`, `https` and `ftp` URLs, and downloads them when you call `render`. Don't pass URLs that come from users without checking them. See [Security](Security).

**Can I write the document straight to a file?**
`render` returns a `StringIO`. Write its `.string` with `File.binwrite("out.docx", word.render.string)`.

**Can docxi stream very large documents?**
No. The whole document is built in memory when you call `render`. See [Deployment](Deployment).

**Can I set right-to-left paragraphs?**
There are no right-to-left options in 0.1.0. Text is written as you pass it.

**Where are the known issues listed?**
On the [Roadmap](Roadmap) and in the README's "Limitations and known issues" section.

## Related pages

- [Troubleshooting](Troubleshooting)
- [Features](Features)
- [Getting Started](Getting-Started)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
