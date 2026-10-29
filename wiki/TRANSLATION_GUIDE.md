# Wiki translation guide

The English pages in `wiki/` are the reference. Translations live in `wiki/i18n/<code>/` with the same file names. The languages are:

`zh-CN`, `zh-TW`, `es`, `ar`, `fr`, `pt-BR`, `de`, `hi`, `ja`, `ko`, `ru`, `id`, `tr`, `it` and `vi`.

The README translations in `docs/i18n/` follow the same rules. See `docs/i18n/TRANSLATION_GUIDE.md`.

## What a translated page contains

- Line 1 is a comment naming the English source:

  ```html
  <!-- Translated from wiki/<Page>.md (git blob <hash>). See wiki/TRANSLATION_GUIDE.md. -->
  ```

  `<hash>` is `git hash-object wiki/<Page>.md` at the time of translation.
- The same sections in the same order as the English page.
- Every `##` and `###` heading starts with an anchor using the English heading's id, for example `## <a id="installation"></a>安装`. Links such as `Troubleshooting#installation` then work in every language.
- The same code blocks, character for character. Only comments inside code may be translated.
- Links to other wiki pages use the English page names (`[安装](Installation)`). The publisher points them at the same language.
- The same external links. Only Home and Credits link to the Phaedra Solutions website.

## What to translate and what to keep

Translate prose, headings, table text, list items, link text, Mermaid labels and comments in code.

Keep unchanged:

- code, commands, environment variables, class and method names, option names, file paths, URLs and version numbers
- product names: docxi, docxer, Ruby, Rails, Microsoft Word, LibreOffice, Google Docs, nokogiri, rubyzip, Bundler, GitHub
- the name Phaedra Solutions

Keep the claims identical, including every known issue and limitation.

Keep headings distinct. If two English headings translate to the same words, reword one of them.

Use the terms developers in that language search for. Use the Microsoft Word terms of that language for headers, footers, shading, mail merge and similar features. When a term has no reliable translation, keep the English term and explain it briefly.

Refer to the people named in the wiki without gendered pronouns or gendered word forms. Use the name, or a neutral construction such as "maintenance of docxi is handled by…".

In Chinese, Japanese and Korean, keep punctuation outside bold markers, as in `**依赖**。`. GitHub doesn't end bold text when `**` follows punctuation and is directly followed by a letter, so `**依赖。**安装` shows the asterisks.

Arabic pages follow the same rules as the Arabic README:

- wrap the page in `<div dir="rtl">`
- start each paragraph, heading, list item and table row with an Arabic word
- wrap inline code in `<span dir="ltr">…</span>`
- never start a code block with an Arabic comment

## Completeness

A language is published only when all of its pages exist. Until then, nothing in that language is published or linked. Never add a page that is only partly translated, or one that still contains English prose.

## When an English page changes

1. Run `ruby wiki/scripts/publish.rb --check`. It lists every translated page whose line-1 hash no longer matches its English page.
2. To see what changed, run `git diff <old-hash> HEAD:wiki/<Page>.md`.
3. Update each affected translation, and its line-1 hash.
4. Update `wiki/i18n/translation-status.json`. If a language can't be updated yet, set its `status` to `outdated`.
5. Run `--check` and `ruby docs/scripts/check_docs.rb` again until neither reports a problem. The second also checks that each translation keeps the English headings, code and links, and that `translation-status.json` matches the files. The Docs workflow runs both on every pull request.

## Status file

`wiki/i18n/translation-status.json` records, for each language:

- whether it's complete
- the English and translated hash of each page
- its status

The status values are:

| Value | Meaning |
|---|---|
| `current` | Every page matches the English page it was translated from |
| `outdated` | One or more English pages changed after translation |
