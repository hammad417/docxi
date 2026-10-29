# Documentation checks

What the documentation checks cover and how to run them. The Docs workflow (`.github/workflows/docs.yml`) runs all three on pull requests and on pushes to `master` that touch Markdown, `docs/`, `wiki/` or the lint config. It only reads the repository, needs no secrets and publishes nothing.

```sh
ruby docs/scripts/check_docs.rb
ruby wiki/scripts/publish.rb --check
npx markdownlint-cli2
```

The two Ruby scripts use only the standard library, so they run without `bundle install`.

## check_docs.rb

- Required files exist: the README, licence, changelog, example, translation guides and status files, and the wiki source.
- Every language has a README translation and all 19 wiki pages, and the language lists agree everywhere: the README selector, both status files, the folders and `publish.rb`.
- Each translation's line 1 names the English version it was made from. If English has changed since, the check fails unless the translation is marked `outdated`. The hashes in both status files must match the files on disk.
- Each translation has the same headings, code blocks and links as the English page. Comments inside code may differ.
- Relative links and `#anchors` resolve. Links to the wiki name a page that `publish.rb` will produce, in the reader's language.
- Every page has one H1 and doesn't skip heading levels, and there's no placeholder text.
- Arabic pages are wrapped in `<div dir="rtl">`, with inline code in `<span dir="ltr">`.
- "Phaedra Solutions" is spelled one way and linked one way: once from each README, and only from the wiki's Home and Credits pages.

Warnings, such as a translation marked `outdated`, are printed but don't fail the run.

## publish.rb --check

Builds the wiki into a temporary folder and checks it as GitHub will serve it: page names GitHub accepts, no two pages colliding, every link between pages and every anchor resolving, links to repository files pointing at real files, and balanced code fences. It also lists stale translations. Nothing in the repository changes.

## markdownlint

Uses `.markdownlint-cli2.jsonc`. Line length, the compact `|---|` table style and bold labels used as small headings are allowed. So is starting a file with something other than a title, because Arabic pages open with `<div dir="rtl">`; `check_docs.rb` checks titles instead. The `a`, `div` and `span` tags the translations need are allowed.

## What they don't catch

None of the checks can tell whether a translation reads naturally, or how a page looks once rendered. For changes to tables, the Mermaid diagram or Arabic pages, open GitHub's Markdown preview as well.
