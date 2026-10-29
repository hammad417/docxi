# Publishing the docxi wiki

GitHub stores a wiki in its own git repository, `https://github.com/hammad417/docxi.wiki.git`, separate from the code. Pages are served by file name only (`/wiki/Getting-Started`), and folders in the wiki repository don't appear in URLs, so two files with the same name in different folders would collide. The publisher solves this by writing every page flat, with a language suffix for translations.

Publishing is always a manual step. `wiki/scripts/publish.rb` writes files into a folder you choose and never runs `git commit` or `git push`.

## Naming convention

| Source file | Published file | URL |
|---|---|---|
| `wiki/Home.md` | `Home.md` | `/wiki` |
| `wiki/Getting-Started.md` | `Getting-Started.md` | `/wiki/Getting-Started` |
| `wiki/i18n/ja/Getting-Started.md` | `Getting-Started-ja.md` | `/wiki/Getting-Started-ja` |
| `wiki/i18n/pt-BR/FAQ.md` | `FAQ-pt-BR.md` | `/wiki/FAQ-pt-BR` |
| `wiki/_Sidebar.md`, `wiki/_Footer.md` | `_Sidebar.md`, `_Footer.md` | Shown on every page |

Names only contain letters, digits and hyphens. The publisher fails if two pages would collide, ignoring case.

## First-time setup

1. In the repository settings, make sure **Wikis** is enabled.
2. Open the Wiki tab and save any first page in the web interface. GitHub creates the wiki's git repository when its first page is saved, and before that it can't be cloned. The publisher overwrites this page.
3. Clone the wiki next to the code:

   ```sh
   git clone https://github.com/hammad417/docxi.wiki.git ../docxi.wiki
   ```

## Publishing

1. Check the source:

   ```sh
   ruby wiki/scripts/publish.rb --check
   ruby docs/scripts/check_docs.rb
   ```

   Fix any broken links, stale translations or other problems they report.
2. Write the pages into the wiki clone:

   ```sh
   ruby wiki/scripts/publish.rb --out ../docxi.wiki
   ```

   Add `--prune` to delete pages in the clone that the source no longer produces. Without it, such pages are listed but left alone.
3. Look at the result:

   ```sh
   git -C ../docxi.wiki status
   git -C ../docxi.wiki diff
   ```

4. Commit and push it yourself:

   ```sh
   git -C ../docxi.wiki add -A
   git -C ../docxi.wiki commit -m "Update wiki"
   git -C ../docxi.wiki push
   ```

5. Open the wiki on GitHub. Check the sidebar, the language switcher at the top of a few pages, and a translated page in each published language.

## What the publisher does

- Copies each English page and adds a language switcher at the top that lists every published language.
- Publishes each complete language with suffixed names, and points its links at the same language. Incomplete languages are skipped and reported.
- Replaces `<!-- languages -->` in `_Sidebar.md` with links to each published language's Home page, and copies `_Footer.md`.
- Validates the result. Every page link and anchor must resolve, every `https://github.com/hammad417/docxi/blob/master/...` link must point to a file in the repository, and code fences must be balanced. It exits with an error if anything fails.

## Updating the sidebar or footer

Edit `wiki/_Sidebar.md` or `wiki/_Footer.md`, keep the `<!-- languages -->` line in the sidebar, then publish as above. Both are shared by all languages, so keep them short and link only to English pages. The language links are generated.

## Undoing a publish

The wiki is a normal git repository. Revert the commit in `../docxi.wiki` and push.

## Repository links

Wiki pages link to repository files on `master`, for example the translated READMEs in `docs/i18n/`. Those links only work once the files are on `master`, so merge documentation changes before publishing the wiki pages that link to them.

The links also go the other way: the README and its translations link to the wiki's Home and Usage Examples pages, in each reader's language. While the wiki has no pages, GitHub redirects those links to the repository's main page. Publish the wiki soon after merging.
