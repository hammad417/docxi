<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API リファレンス：クラスとメソッド

docxi は Ruby のライブラリなので、このページでは Ruby の API を一覧にしています。HTTP API はありません。各メソッドのオプションは[機能](Features)で、クラスについては[アーキテクチャ](Architecture)で説明しています。

## <a id="docxidocument"></a>Docxi::Document

| メソッド | 戻り値 | 備考 |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` はページの余白です。[設定](Configuration)を参照してください |
| `#document` | `Docxi::Word::Document` | コンテンツを追加する本文 |
| `#render` | `StringIO` | 完全な `.docx` パッケージ。`.string` でバイト列を取得できます |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document（`word.document`）

| メソッド | 戻り値 | 追加するもの |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | 段落 |
| `text(string, options = {})` | Paragraph | テキストランが 1 つの段落 |
| `br(page: true)` | Break | 改ページ |
| `tab(page: true)` | Tab | タブを 1 つ含む段落 |
| `table(options = {}) { \|table\| }` | Table | 表 |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | 目次の項目 |
| `frame(text, options = {})` | Paragraph | テキストボックスを含む段落 |
| `add_media(io_path_or_url, options = {})` | Media | `image` で配置する画像 |
| `add_hyperlink(url, options = {})` | String | `hyperlink` 用のリレーションシップ ID |
| `add_header(header)`、`add_first_header(header)` | Header | 2 ページ目以降のヘッダーと、1 ページ目のヘッダー |
| `add_footer(footer)`、`add_first_footer(footer)` | Footer | 2 ページ目以降のフッターと、1 ページ目のフッター |

## <a id="paragraph-docxiwordcontentsparagraph"></a>段落（`Docxi::Word::Contents::Paragraph`）

| メソッド | 追加するもの |
|---|---|
| `text(string, options = {})` | テキストラン。段落のオプションがデフォルトになります |
| `br(times: 1)` | 改行 |
| `tab(times: 1)` | タブ文字 |
| `image(media, options = {})` | インライン画像または配置を指定した画像 |
| `hyperlink(text, url, id)` | `add_hyperlink` が返したリレーションシップの `id` を使うリンク |
| `frame(options = {}) { \|frame\| }` | テキストボックス |

## <a id="tables"></a>表

| メソッド | 追加するもの |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | 行。表のオプションがデフォルトになります |
| `TableRow#tc(options = {}) { \|tc\| }` | セル。行のオプションがデフォルトになります |
| `TableCell#text(string, options = {})` | 段落。セルのオプションがデフォルトになります |
| `TableCell#p(options = {}) { \|p\| }` | 独自のオプションを持つ段落 |
| `TableCell#image(media, width:, height:)` | 画像 |
| `TableCell#page_numbers` | ページ番号フィールド |

## <a id="table-of-contents"></a>目次

| メソッド | 追加するもの |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | 目次の項目 1 つ |

## <a id="headers-and-footers"></a>ヘッダーとフッター

| クラス | メソッド |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`、`br`、`table`、`image`、`page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`、`br`、`table`、`page_numbers(options = {})` |

## <a id="frames"></a>フレーム

`Docxi::Word::Contents::Frame` のブロック内では `text`、`br`、`tab`、`image`、`frame` を使えます。

## <a id="constants-and-helpers"></a>定数とヘルパー

| 名前 | 意味 |
|---|---|
| `Docxi::VERSION` | インストールされているバージョン（例：`"0.1.0"`） |
| `Docxi.write_entry(zip, name, data)` | サイズを先に設定したうえで、パーツを 1 つアーカイブに書き込みます。すべてのパーツが内部で使っています |
| `Docxi::Util.blank?`、`Docxi::Util.present?` | ActiveSupport の `blank?` と同じ規則を持つ内部ヘルパー |

## <a id="related-pages"></a>関連ページ

- [機能](Features)
- [使用例](Usage-Examples)
- [アーキテクチャ](Architecture)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
