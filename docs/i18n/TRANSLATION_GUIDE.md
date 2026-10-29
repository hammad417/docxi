# Translation guide

The English [README.md](../../README.md) at the repository root is the source for every translation. The translations live in this folder as `README.<code>.md`, and [translation-status.json](translation-status.json) records which version of the English README each one was made from.

## What every translation contains

- Line 1 is a comment naming the English source: `<!-- Translated from README.md (git blob <hash>). ... -->`. The hash is `git hash-object README.md` at the time of translation.
- The language selector with all 16 languages, English first, in the same order as in the English README.
- A short note, in the target language, saying the English README is authoritative when the two differ.
- The same sections in the same order as the English README.
- Every `##` and `###` heading starts with an anchor that uses the English section id, for example `## <a id="overview"></a>Обзор`. The contents list and all in-page links point to these ids, so they work in every language and match the English anchors.
- The same code blocks, character for character, except that comments inside code may be translated.
- The same links. Paths are relative to `docs/i18n/`, so files at the repository root are `../../LICENSE`, `../../CHANGELOG.md` and so on, and the other translations are `README.<code>.md`.

## What to translate and what to keep

Translate:

- prose, headings, list items and table text
- FAQ questions and answers, and troubleshooting entries
- badge alt text
- Mermaid node labels
- comments inside code blocks

Keep exactly as in English:

- code, commands, class and method names, option names such as `space: true`, file paths, URLs and version numbers
- the product and project names: docxi, docxer, Ruby, Rails, Ruby on Rails, Microsoft Word, LibreOffice, Google Docs, nokogiri, rubyzip, Bundler, GitHub
- the brand name Phaedra Solutions and its URL, `https://www.phaedrasolutions.com/`
- the maintainer's name and the original author's name. A transliteration may be added in brackets, as the Russian translation does.

Keep the claims identical. A translation must not add, drop or soften a feature, a limitation, a supported version or the attribution.

Keep headings distinct. When two English headings translate to the same words, reword the second one. "Contents" and "Table of contents" are both 目次 in Japanese, for example, so the usage example is titled 目次の追加 (adding a table of contents).

Use the words developers in that language actually use, including the terms Microsoft Word uses in that language for headers, footers, shading, mail merge and similar features. When a term has no reliable translation, keep the English term and explain it in a few words.

## Arabic

- Keep the whole file inside `<div dir="rtl">` … `</div>`, with a blank line after the opening tag and before the closing tag.
- GitHub sets each paragraph's direction from its first letter. Start every paragraph, heading, list item and table row with an Arabic word rather than with "docxi" or a code span.
- Wrap every inline code span in `<span dir="ltr">…</span>`. Without it, characters such as `.`, `>`, `\` and `/` are reordered inside right-to-left text, and `.docx` shows as `docx.`.
- Code blocks must start with code, not with an Arabic comment. Otherwise GitHub renders the whole block right to left. Put such a comment after the first line of code, as the configuration example does.

## When README.md changes

1. Commit the change to `README.md`. Then see what changed since the translations were made:

   ```sh
   git diff <source_blob> HEAD:README.md
   ```

   `<source_blob>` is the value in `translation-status.json`.
2. Apply the same change to each translation. If you can't translate it right away, set that language's `status` to `outdated` instead of copying English text into it.
3. Update line 1 of each updated translation with the new hash from `git hash-object README.md`.
4. Update `translation-status.json`:
   - `source.blob` and each entry's `source_blob` get the new README hash.
   - Each updated file gets a new `translation_blob` from `git hash-object docs/i18n/README.<code>.md`.
   - Set `status` to `current` for each language you updated.
5. Run the checks below.

## Checks before merging

Run:

```sh
ruby docs/scripts/check_docs.rb
```

It fails if a translation:

- names an older `README.md` on line 1 without being marked `outdated`, or doesn't match `translation-status.json`
- has different headings, code blocks or links from the English README. Comments inside code may differ
- has an `](#...)` link without a matching anchor, or a relative link that doesn't resolve from `docs/i18n/`
- has a different language selector, or links to a wiki page in another language
- is Arabic and isn't wrapped in `<div dir="rtl">`, or has inline code outside `<span dir="ltr">`
- doesn't link to the Phaedra Solutions website exactly once

The Docs workflow (`.github/workflows/docs.yml`) runs it, the wiki check and markdownlint on every pull request that changes documentation.

It can't see how the page looks, so also open the translation in GitHub's Markdown preview and look at the tables, the Mermaid diagram and, for Arabic, the text direction.

## Status values

| Value | Meaning |
|---|---|
| `current` | The translation matches the `README.md` named in its `source_blob` |
| `outdated` | `README.md` changed after this translation's `source_blob`, and the translation hasn't been updated yet |

## Adding a language

1. Copy an existing translation as `README.<code>.md` and translate it, following the rules above.
2. Add the language to the selector in `README.md` and in every file in this folder, keeping English first.
3. Add it to [README.md](README.md) in this folder and to `translation-status.json`.
