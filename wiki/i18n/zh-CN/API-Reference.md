<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API 参考：类和方法

docxi 是一个 Ruby 库，所以本页列出的是它的 Ruby API，它没有 HTTP API。各方法的选项在[功能](Features)中说明，各个类在[架构](Architecture)中解释。

## <a id="docxidocument"></a>Docxi::Document

| 方法 | 返回值 | 说明 |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` 是页边距。参见[配置](Configuration) |
| `#document` | `Docxi::Word::Document` | 用于添加内容的正文 |
| `#render` | `StringIO` | 完整的 `.docx` 包。`.string` 返回字节内容 |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document（`word.document`）

| 方法 | 返回值 | 添加的内容 |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | 一个段落 |
| `text(string, options = {})` | Paragraph | 一个只包含一个文本片段的段落 |
| `br(page: true)` | Break | 一个分页符 |
| `tab(page: true)` | Tab | 一个只包含一个制表符的段落 |
| `table(options = {}) { \|table\| }` | Table | 一个表格 |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | 目录行 |
| `frame(text, options = {})` | Paragraph | 一个包含文本框的段落 |
| `add_media(io_path_or_url, options = {})` | Media | 一张用 `image` 放置的图片 |
| `add_hyperlink(url, options = {})` | String | 供 `hyperlink` 使用的关系 ID |
| `add_header(header)`、`add_first_header(header)` | Header | 第 2 页起的页眉，以及第 1 页的页眉 |
| `add_footer(footer)`、`add_first_footer(footer)` | Footer | 第 2 页起的页脚，以及第 1 页的页脚 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>段落（`Docxi::Word::Contents::Paragraph`）

| 方法 | 添加的内容 |
|---|---|
| `text(string, options = {})` | 一个文本片段。段落的选项是它的默认值 |
| `br(times: 1)` | 换行 |
| `tab(times: 1)` | 制表符 |
| `image(media, options = {})` | 一张内联图片或定位图片 |
| `hyperlink(text, url, id)` | 一个使用 `add_hyperlink` 返回的关系 `id` 的链接 |
| `frame(options = {}) { \|frame\| }` | 一个文本框 |

## <a id="tables"></a>表格

| 方法 | 添加的内容 |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | 一行。表格的选项是它的默认值 |
| `TableRow#tc(options = {}) { \|tc\| }` | 一个单元格。行的选项是它的默认值 |
| `TableCell#text(string, options = {})` | 一个段落。单元格的选项是它的默认值 |
| `TableCell#p(options = {}) { \|p\| }` | 一个使用自身选项的段落 |
| `TableCell#image(media, width:, height:)` | 一张图片 |
| `TableCell#page_numbers` | 一个页码域 |

## <a id="table-of-contents"></a>目录

| 方法 | 添加的内容 |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | 一行目录 |

## <a id="headers-and-footers"></a>页眉和页脚

| 类 | 方法 |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`、`br`、`table`、`image`、`page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`、`br`、`table`、`page_numbers(options = {})` |

## <a id="frames"></a>框架

在 `Docxi::Word::Contents::Frame` 的块中可以调用 `text`、`br`、`tab`、`image` 和 `frame`。

## <a id="constants-and-helpers"></a>常量和辅助方法

| 名称 | 含义 |
|---|---|
| `Docxi::VERSION` | 已安装的版本，例如 `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | 把一个部件写入压缩包，并预先设置其大小。所有部件在内部都使用它 |
| `Docxi::Util.blank?`、`Docxi::Util.present?` | 内部辅助方法，遵循 ActiveSupport 的 `blank?` 规则 |

## <a id="related-pages"></a>相关页面

- [功能](Features)
- [使用示例](Usage-Examples)
- [架构](Architecture)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
