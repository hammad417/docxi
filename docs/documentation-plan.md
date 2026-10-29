# Documentation

How the docxi documentation is organised and how to keep it up to date.

## Where things live

| Path | What it is |
|---|---|
| `README.md` | The main page: what docxi does and doesn't do, installation, examples, configuration, troubleshooting, FAQ |
| `docs/i18n/README.<code>.md` | The README in 15 other languages |
| `wiki/` | Source of the GitHub wiki, in English and the same 15 languages, with `scripts/publish.rb` to publish it |
| `CHANGELOG.md` | Release notes |
| `examples/full_document.rb` | One script that uses every feature |
| `docs/scripts/check_docs.rb` | Checks the translations, links, headings and status files |
| `docs/repository-analysis.md` | Notes on how the code is organised and how it behaves |
| `docs/seo-keyword-strategy.md`, `docs/SEO_AUDIT.md`, `docs/GITHUB_DISCOVERABILITY.md`, `docs/ASO_OR_PACKAGE_SEO.md` | How docxi is described on GitHub, RubyGems and in search |
| `docs/FINAL_DOCUMENTATION_AUDIT.md` | What the documentation checks cover and how to run them |

The README is kept short. The full option reference, longer guides, internals and the release process are in the wiki.

## Languages

English is the source. The translations are zh-CN, zh-TW, es, ar, fr, pt-BR, de, hi, ja, ko, ru, id, tr, it and vi, for both the README and the wiki. Each translation's title uses the phrase a developer in that language would search for, not a word-for-word translation of the English title.

Every README starts with the same language selector, English first. The rules for translating, and what to do when the English text changes, are in `docs/i18n/TRANSLATION_GUIDE.md` and `wiki/TRANSLATION_GUIDE.md`.

## Wiki

The 19 wiki pages are written in `wiki/` and published to the GitHub wiki with `wiki/scripts/publish.rb`, which gives translated pages a language suffix such as `Getting-Started-ja`. `wiki/wiki-publishing-guide.md` has the steps. The README and its translations link to the wiki's Home and Usage Examples pages in each reader's language.

## Keeping it current

- When behaviour changes, update the README section and the wiki page in the same pull request, and the translations with them. If a translation can't be updated straight away, mark it `outdated` in its status file.
- On every release, update `CHANGELOG.md`, the wiki's Home and Changelog pages, and anything else that names a version.
- Only describe what the code does. Known issues belong on the wiki's Roadmap page.
- Before opening a pull request that touches documentation, run:

  ```sh
  ruby docs/scripts/check_docs.rb
  ruby wiki/scripts/publish.rb --check
  npx markdownlint-cli2
  ```

  The Docs workflow runs the same three on every pull request.
