# docxi wiki source

This folder is the source of the [docxi GitHub wiki](https://github.com/hammad417/docxi/wiki). Pages are written here, in the main repository, and published to the wiki's separate git repository with `wiki/scripts/publish.rb`. This file explains how the pieces fit together.

## Layout

| Path | Purpose |
|---|---|
| `wiki/*.md` | English pages. English is the reference version |
| `wiki/_Sidebar.md`, `wiki/_Footer.md` | Navigation shown on every wiki page |
| `wiki/i18n/<code>/*.md` | Translated pages, one folder per language, with the same file names as the English pages |
| `wiki/i18n/translation-status.json` | Completion status of each language, with the version of each English page it was translated from |
| `wiki/scripts/publish.rb` | Builds the published pages and checks every link |
| `wiki/TRANSLATION_GUIDE.md` | How to translate pages and keep translations current |
| `wiki/wiki-publishing-guide.md` | How to publish to GitHub |

This README and the two guides are not published as wiki pages.

## Writing a page

- Each page is one Markdown file named after the page, such as `Getting-Started.md`. GitHub shows the name with spaces in place of hyphens.
- Start with a descriptive `#` title and a short introduction. End with a "Related pages" list and the line linking back to the wiki Home and the README.
- Link to other pages by page name, without `.md`: `[Installation](Installation)` or `[Troubleshooting](Troubleshooting#installation)`.
- Link to repository files with full URLs on `master`, such as `https://github.com/hammad417/docxi/blob/master/CHANGELOG.md`.
- Ruby examples must run as written. A page's Ruby blocks are run in order, so later blocks can build on earlier ones.
- Only describe what the code does. Known issues go on the [Roadmap](Roadmap.md) page.
- Link to the Phaedra Solutions website only from Home and Credits.

## Navigation

- `_Sidebar.md` and `_Footer.md` appear on every page, in every language.
- The `<!-- languages -->` line in `_Sidebar.md` is replaced at publish time with links to the Home page of every published language.
- Each published page starts with a language switcher, added at publish time, that links to the same page in every published language.
- A language is published only when all of its pages exist, so the switcher and sidebar never link to a missing page.

## Published page names

GitHub serves wiki pages by file name only, so translated pages get a language suffix:

| Source | Published as |
|---|---|
| `wiki/Getting-Started.md` | `Getting-Started` |
| `wiki/i18n/zh-CN/Getting-Started.md` | `Getting-Started-zh-CN` |
| `wiki/i18n/pt-BR/Home.md` | `Home-pt-BR` |

Inside a translation, links are written with English page names, and the publisher points them at the same language. `[Installation](Installation)` in a German page becomes `Installation-de`.

## Checking the wiki

```sh
ruby wiki/scripts/publish.rb --check
```

This builds the wiki into a temporary folder and reports:

- which languages are complete
- which translations are stale
- every broken page link, anchor or repository link

It changes nothing in the repository and needs no network access. Run it before every publish and in any pull request that touches `wiki/`.

`ruby docs/scripts/check_docs.rb` adds the checks that compare each translation with its English page: headings, code blocks and links, the line-1 hash, and `wiki/i18n/translation-status.json`. It also checks that only Home and Credits link to the Phaedra Solutions website. The Docs workflow (`.github/workflows/docs.yml`) runs both scripts and markdownlint on every pull request that changes documentation.

## Stale translations

Line 1 of every translated page names the git blob hash of the English page it was translated from. `--check` compares it with the current English page and lists every translation that is out of date. See [TRANSLATION_GUIDE.md](TRANSLATION_GUIDE.md).
