# GitHub settings

The repository description, topics and website aren't stored in files, so they're noted here along with the values to use.

## Description

GitHub puts the description in the repository page's title and meta description, and allows up to 350 characters. This one is 143, with no final period:

> Generate Microsoft Word (.docx) documents with Ruby: paragraphs, tables, images, headers, footers and page numbers. Works with or without Rails

The old description, "Help to generate formatted docx template.", talked about templates, which docxi doesn't do.

## Topics

Up to 20 topics, lowercase letters, numbers and hyphens, 50 characters each. The counts are how many public repositories used each topic in October 2026. They show how crowded a topic page is, not how often people search for it.

| Topic | Repositories |
|---|---|
| `ruby` | 38,251 |
| `ruby-gem` | 2,464 |
| `rails` | 13,668 |
| `ruby-on-rails` | 11,106 |
| `docx` | 3,482 |
| `docx-generator` | 216 |
| `word` | 2,159 |
| `microsoft-word` | 426 |
| `word-document` | 48 |
| `ooxml` | 406 |
| `office-open-xml` | 42 |
| `openxml` | 241 |
| `wordprocessingml` | 17 |
| `document-generation` | 687 |
| `report-generation` | 536 |

`rubygems` and `office` would also fit. `nokogiri` and `rubyzip` are left out because people browsing those topics are after something else.

## Website

The wiki, `https://github.com/hammad417/docxi/wiki`, once it has pages. Until then the RubyGems page works. The company site doesn't belong here: the field is for the project's own home.

## Social preview

PNG, JPG or GIF under 1 MB, at least 640 × 320, ideally 1280 × 640, set in Settings → General. A good one is the name, the line "Generate Word (.docx) documents with Ruby", and a picture of a real generated page, such as one from `tmp/samples/kitchen_sink.docx`.

## Releases

`release.yml` publishes the gem when a `v*` tag is pushed, but it doesn't create a GitHub Release. To add one, use the version's entry from `CHANGELOG.md`:

```sh
gh release create v0.1.1 --verify-tag --title "docxi 0.1.1" --notes-file notes.md
```

## Applying the settings

```sh
gh repo edit hammad417/docxi \
  --description "Generate Microsoft Word (.docx) documents with Ruby: paragraphs, tables, images, headers, footers and page numbers. Works with or without Rails" \
  --add-topic ruby,ruby-gem,rails,ruby-on-rails,docx,docx-generator,word,microsoft-word,word-document,ooxml,office-open-xml,openxml,wordprocessingml,document-generation,report-generation \
  --homepage "https://github.com/hammad417/docxi/wiki"
```

## Phaedra Solutions

docxi is maintained by Hammad Maqbool, who works at Phaedra Solutions; the gemspec lists his phaedrasolutions.com address. The repository and the gem belong to the `hammad417` account, and the MIT copyright belongs to Timur Gabdrakipov, who wrote docxer. So the docs say docxi is maintained by Hammad Maqbool *at* Phaedra Solutions, not that the company owns or publishes it.

The name is always written "Phaedra Solutions" and the link is always `https://www.phaedrasolutions.com/`. The description of the company, "an AI-first software engineering company that has been building custom software for businesses since 2013", matches the company's own site.

Each README links to the site once, from the company name under "Maintainers". In the wiki only Home and Credits link to it. There are no links in sidebars, footers or code, and the anchor text is always just the company name. `check_docs.rb` keeps it that way.

Because GitHub marks Markdown links `nofollow` and doesn't index small wikis, these links are there to credit the company and let readers find it, not to move search rankings. Going the other way, the company site could list docxi on an open-source page with links to the repository and the RubyGems page.
