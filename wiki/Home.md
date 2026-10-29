# docxi wiki: generate Word (.docx) documents with Ruby

docxi is a Ruby gem that builds Microsoft Word `.docx` files from code. You get paragraphs, formatted text, tables, images, headers and footers, page numbers, hyperlinks and a simple table of contents. It doesn't need Microsoft Word, LibreOffice or a template, and it works in plain Ruby as well as in Ruby on Rails apps.

The [README](https://github.com/hammad417/docxi#readme) is the short introduction. This wiki covers every option, has ready-to-copy recipes, explains how docxi works inside, and helps when something goes wrong.

Current release: **0.1.1**, for Ruby 2.5 or newer. See the [Changelog](Changelog).

## Start here

**New to docxi**

- [Getting Started](Getting-Started): install the gem and create your first document
- [Installation](Installation): requirements, Bundler, and upgrading from 0.0.x
- [Usage Examples](Usage-Examples): reports, invoices, letters and tables from data, ready to copy

**Building documents**

- [Features](Features): every element and the options it accepts
- [Configuration](Configuration): margins, page setup and units
- [API Reference](API-Reference): classes and methods at a glance
- [Best Practices](Best-Practices): habits that avoid the usual surprises

**When something goes wrong**

- [Troubleshooting](Troubleshooting): error messages and documents that don't look right
- [FAQ](FAQ): short answers to common questions
- [Security](Security): reporting problems and handling untrusted input

**Going further**

- [Architecture](Architecture): how Ruby calls become a `.docx` package
- [Integrations](Integrations): Rails controllers, Rack apps and background jobs
- [Deployment](Deployment): running docxi in production, and how releases are published
- [Contributing](Contributing) and [Development Setup](Development-Setup)
- [Roadmap](Roadmap), [Changelog](Changelog) and [Credits](Credits)

## What docxi does, and what it doesn't

docxi creates new documents. It can't open or edit an existing `.docx`, fill in a Word template or run a mail merge. It has no numbered lists, heading styles or automatic table of contents, and pages are always US Letter in portrait. The known issues planned for 0.2.0 are listed on the [Roadmap](Roadmap).

## Credits

docxi is maintained by [Hammad Maqbool](https://github.com/hammad417) at [Phaedra Solutions](https://www.phaedrasolutions.com/). It started in 2016 from [docxer](https://github.com/gtimti/docxer) by Timur Gabdrakipov. See [Credits](Credits) for details.

## Languages

English is the reference version of this wiki. Translations are linked in the sidebar and at the top of every page.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Report an issue](https://github.com/hammad417/docxi/issues)
