<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API 參考：類別和方法

docxi 是一個 Ruby 程式庫，所以本頁列出的是它的 Ruby API，它沒有 HTTP API。各方法的選項在[功能](Features)中說明，各個類別在[架構](Architecture)中解釋。

## <a id="docxidocument"></a>Docxi::Document

| 方法 | 傳回值 | 說明 |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` 是邊界。請參閱[設定](Configuration) |
| `#document` | `Docxi::Word::Document` | 用來加入內容的本文 |
| `#render` | `StringIO` | 完整的 `.docx` 封裝。`.string` 會傳回位元組內容 |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document（`word.document`）

| 方法 | 傳回值 | 加入的內容 |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | 一個段落 |
| `text(string, options = {})` | Paragraph | 一個只含一個文字片段的段落 |
| `br(page: true)` | Break | 一個分頁符號 |
| `tab(page: true)` | Tab | 一個只含一個定位字元的段落 |
| `table(options = {}) { \|table\| }` | Table | 一個表格 |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | 目錄項目 |
| `frame(text, options = {})` | Paragraph | 一個包含文字方塊的段落 |
| `add_media(io_path_or_url, options = {})` | Media | 一張以 `image` 放置的圖片 |
| `add_hyperlink(url, options = {})` | String | 供 `hyperlink` 使用的關聯 ID |
| `add_header(header)`、`add_first_header(header)` | Header | 第 2 頁起的頁首，以及第 1 頁的頁首 |
| `add_footer(footer)`、`add_first_footer(footer)` | Footer | 第 2 頁起的頁尾，以及第 1 頁的頁尾 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>段落（`Docxi::Word::Contents::Paragraph`）

| 方法 | 加入的內容 |
|---|---|
| `text(string, options = {})` | 一個文字片段。段落的選項是它的預設值 |
| `br(times: 1)` | 分行 |
| `tab(times: 1)` | 定位字元 |
| `image(media, options = {})` | 一張內嵌圖片或定位圖片 |
| `hyperlink(text, url, id)` | 一個使用 `add_hyperlink` 傳回之關聯 `id` 的連結 |
| `frame(options = {}) { \|frame\| }` | 一個文字方塊 |

## <a id="tables"></a>表格

| 方法 | 加入的內容 |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | 一列。表格的選項是它的預設值 |
| `TableRow#tc(options = {}) { \|tc\| }` | 一個儲存格。列的選項是它的預設值 |
| `TableCell#text(string, options = {})` | 一個段落。儲存格的選項是它的預設值 |
| `TableCell#p(options = {}) { \|p\| }` | 一個使用自身選項的段落 |
| `TableCell#image(media, width:, height:)` | 一張圖片 |
| `TableCell#page_numbers` | 一個頁碼功能變數 |

## <a id="table-of-contents"></a>目錄

| 方法 | 加入的內容 |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | 一個目錄項目 |

## <a id="headers-and-footers"></a>頁首和頁尾

| 類別 | 方法 |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`、`br`、`table`、`image`、`page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`、`br`、`table`、`page_numbers(options = {})` |

## <a id="frames"></a>框架

在 `Docxi::Word::Contents::Frame` 的區塊中可以呼叫 `text`、`br`、`tab`、`image` 和 `frame`。

## <a id="constants-and-helpers"></a>常數和輔助方法

| 名稱 | 意義 |
|---|---|
| `Docxi::VERSION` | 已安裝的版本，例如 `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | 將一個組件寫入壓縮檔，並事先設定其大小。所有組件在內部都使用它 |
| `Docxi::Util.blank?`、`Docxi::Util.present?` | 內部輔助方法，遵循 ActiveSupport 的 `blank?` 規則 |

## <a id="related-pages"></a>相關頁面

- [功能](Features)
- [使用範例](Usage-Examples)
- [架構](Architecture)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
