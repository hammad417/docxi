# Keyword strategy

The terms below come from what docxi does, the words Ruby developers use for this task, and how related gems describe themselves. There's no search-volume data behind them, so check them against RubyGems downloads or GitHub traffic before leaning on any one of them.

## Keywords

**Primary:** generate Word (.docx) documents with Ruby

**Secondary**

- ruby docx
- create docx in ruby
- ruby word document generator
- rails export to word
- rails docx export
- ruby on rails generate word document
- docx gem

**Language, framework and format terms**

Ruby, Ruby on Rails, RubyGems, Bundler, nokogiri, rubyzip, Microsoft Word, `.docx`, Office Open XML, OOXML, WordprocessingML

**Long-tail developer queries**

- how to generate a word document in ruby on rails
- create a docx file with tables and images in ruby
- add a header and footer with page numbers to a docx in ruby
- export a rails report to a word document
- send a docx file from a rails controller
- docx mime type for send_data

**Problem-solving queries**

- generate word documents on a linux server without microsoft office
- ruby docx library without templates
- word export for rails 7 or rails 8

**Installation and troubleshooting queries**

- gem install docxi
- docxi rails 8
- bundler could not find compatible versions for gem "nokogiri" docxi
- rubyzip 3 docx won't open in google docs
- rubyzip zip64 docx libreoffice
- undefined method present? docxi
- application/vnd.openxmlformats-officedocument.wordprocessingml.document

**Industry terms:** document generation, document automation, report export, Office Open XML

## Where the keywords appear in the README

| Location | Wording |
|---|---|
| Title | "create Word (.docx) documents with Ruby" |
| Opening paragraph | Ruby gem, Microsoft Word `.docx`, Office Open XML, Ruby on Rails |
| Features | tables, images, headers and footers, page numbers, hyperlinks, table of contents |
| Rails example | `send_data`, the `.docx` MIME type |
| Troubleshooting | the exact error messages people search for: the nokogiri and rubyzip conflicts, `present?`, Zip64 in Google Docs, `Errno::ENOENT` |
| FAQ | "How do I generate a Word document in Ruby on Rails?", "Do I need Microsoft Word or LibreOffice installed?", "Can docxi edit an existing .docx or fill in a Word template?" |

Each term appears where it describes something real, and only once or twice. Headings stay descriptive rather than keyword lists.

## Translated READMEs

The keyword matrix for all 16 languages is in [SEO_AUDIT.md](SEO_AUDIT.md#terms-by-language).

The 15 translations in `docs/i18n/` target the same search intent in each language, rather than translating the English keywords word for word.

- The title, opening paragraph, feature list, use cases and FAQ questions use the phrasing a developer in that language would type. For example, the Japanese title uses "Ruby で Word（.docx）文書を生成", and the Spanish FAQ asks "¿Cómo genero un documento de Word en Ruby on Rails?".
- Feature names use the terms Microsoft Word itself uses in that language, which are also what people search for: 页眉/页脚, 머리글/바닥글, Kopf- und Fußzeilen, en-têtes et pieds de page, колонтитулы, and so on.
- Badge alt text and the anchor text of internal links are translated. Link targets and the explicit section ids stay the same in every language.
- Product names, code, error messages and commands stay in English, because that is how people search for them.
- Every translation makes exactly the claims the English README makes, and keeps the same single attribution section.

## Repository and gem metadata

The recommendations for the GitHub description, topics and website are in [GITHUB_DISCOVERABILITY.md](GITHUB_DISCOVERABILITY.md), and those for the gem summary, description and documentation link are in [ASO_OR_PACKAGE_SEO.md](ASO_OR_PACKAGE_SEO.md).

## Related gems

These solve neighbouring problems. Knowing them helps docxi describe itself accurately and point people elsewhere when docxi isn't the right tool.

| Gem | Approach |
|---|---|
| caracal | Builds new `.docx` documents from a Ruby DSL, the same approach as docxi |
| sablon | Fills `.docx` templates that contain merge fields |
| docx | Reads and edits existing `.docx` files |
| htmltoword | Converts HTML into `.docx` |

Next to these, docxi:

- builds new documents from code
- two runtime dependencies
- no Rails dependency
- CI-tested from Ruby 2.5 to 4.0 and Rails 5.2 to 8.1
- output kept stable across releases by golden-file tests

It shouldn't claim to be more complete than caracal, or to do what sablon or docx do.

## What to avoid

- Repeating "docx" or "Word" in every heading or sentence
- Describing features docxi doesn't have, such as templates, editing existing files, automatic tables of contents or numbered lists
- Superlatives such as "best" or "fastest" with nothing to back them up
- Translating keywords word for word in the localized READMEs. Each translation should use the phrase developers search for in that language

## Measuring

- RubyGems download counts per version: `https://rubygems.org/api/v1/versions/docxi.json`
- GitHub Insights → Traffic (visible to the maintainer): views, unique visitors, referring sites and popular content for the last 14 days
- Compare the two weeks before and after changing the GitHub description, topics or README, and write the numbers down. GitHub only keeps 14 days of traffic.
