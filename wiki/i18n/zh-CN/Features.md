<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 功能：段落、文本、表格、图片、页眉和页脚

本页列出 docxi 可以添加到 Word 文档中的每种元素，以及每种元素支持的选项。示例都基于以下设置：

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

不需要的选项可以省略。颜色是不带 `#` 的十六进制值，例如 `"8B0000"`。单位和页面设置参见[配置](Configuration)。

## <a id="paragraphs"></a>段落

`document.p(options) { |p| ... }` 添加一个段落。

| 选项 | 取值 | 效果 |
|---|---|---|
| `align` | `"left"`、`"center"`、`"right"`、`"both"` | 对齐方式。`"both"` 表示两端对齐 |
| `fill` | 十六进制颜色 | 段落及其文本背后的底纹 |
| `left`、`right` | 二十分之一磅 | 缩进。两者要同时设置，只设置其中一个会被忽略 |
| `ul` | `true` | 将段落设为项目符号列表项 |
| `bottom` | `true` | 在段落下方绘制边框 |

`size:`、`bold:`、`color:` 等文本选项也可以传给 `p`，它们会成为段落中每个文本片段的默认值。

在块（block）中可以调用 `text`、`br`、`tab`、`image`、`hyperlink` 和 `frame`。`document.text(string, options)` 是只包含一个文本片段的段落的简写。

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>文本

`p.text(string, options)` 添加一个文本片段（run）。

| 选项 | 取值 | 效果 |
|---|---|---|
| `bold` | `true` | 粗体 |
| `italic` | `true` | 斜体 |
| `underline` | `true` | 单下划线 |
| `color` | 十六进制颜色 | 文本颜色 |
| `size` | 磅 | 字号 |
| `font` | 字体名称 | 字体，例如 `"Georgia"` |
| `fill` | 十六进制颜色 | 文本背后的底纹 |
| `space` | `true` | 保留文本开头和结尾的空格 |
| `br` | `true` | 在文本后添加换行 |

`nil` 字符串会被跳过。

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>换行、分页和制表符

- `p.br` 或 `p.br(times: 3)` 在段落内添加换行。
- `document.br(page: true)` 开始新的一页。
- `p.tab` 或 `p.tab(times: 2)` 在段落内添加制表符。
- `document.tab(page: true)` 添加一个只包含一个制表符的段落。

不要在不带 `page: true` 的情况下调用 `document.br` 或 `document.tab`。它们会把文本片段直接放进文档正文，这不是有效的 WordprocessingML。

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>表格

`document.table(options)` 接受一个块，在块中用 `table.tr(options)` 添加行，用 `tr.tc(options)` 添加单元格。

**表格选项**

| 选项 | 取值 | 效果 |
|---|---|---|
| `columns_width` | 数字数组 | 列宽，乘以 14.92 换算为二十分之一磅。也是单元格的默认宽度 |
| `height` | 二十分之一磅 | 每一行的最小高度 |
| `iwidth` | 二十分之一磅 | 表格相对左边距的缩进 |
| `style` | 样式 ID | 表格样式。docxi 定义了 `TableNormal` 和 `TableCustom` |
| `width` | 数字 | 表格的指定宽度，类型写为 `auto`。未设置 `columns_width` 时，每个单元格也会使用这个宽度，并乘以 14.92 |
| `align` | `"left"`、`"center"`、`"right"` | 表格对齐方式。用 `tc.text` 添加的文本也按同样方式对齐 |
| `padding` | `true` | 在单元格内添加固定的内边距 |
| `borders` | 哈希 | 已知问题：会应用到每个单元格，而不是表格的外框 |

**行选项**：`height:` 设置行的最小高度。传给 `tr` 的任何单元格选项都会作用于该行的每个单元格。表格选项也以同样的方式传递给每一行和每个单元格。

**单元格选项**

| 选项 | 取值 | 效果 |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | 边框宽度，单位为八分之一磅。`nil` 表示该边没有边框 |
| `fill` | 十六进制颜色 | 单元格底纹。没有设置填充的单元格为白色（`"FFFFFF"`） |
| `valign` | `"top"`、`"center"`、`"bottom"` | 垂直对齐 |
| `colspan` | 整数 | 单元格跨越的列数 |
| `rowspan` | `true` | 开始一个垂直合并 |
| `merged` | `true` | 延续上方单元格的垂直合并，其内容会被忽略 |
| `width` | 数字 | 单元格宽度，乘以 14.92 |

**单元格内容**

- `tc.text(string, options)`：单元格和行的选项（例如 `bold:` 或 `align:`）同样作用于文本。
- `tc.p(options) { |p| ... }`：一个只使用自身选项的段落。
- `tc.image(media, width:, height:)`：一张图片。尺寸问题参见下文的[图片](#images)。
- `tc.page_numbers`：一个页码域。已知问题：它前面会带有文本“GlobalOptions”。

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>图片

`media = document.add_media(source)` 注册一张图片。`source` 可以是：

- IO 对象，例如 `File` 或 `StringIO`
- 文件路径
- `http`、`https` 或 `ftp` URL

图片在调用 `render` 时读取，并且始终以 PNG 部件存储，所以 PNG 文件是稳妥的选择。

**内联图片**：`p.image(media, options)` 把图片放在文本行中。

| 选项 | 取值 | 效果 |
|---|---|---|
| `style` | CSS 属性的哈希 | 尺寸，例如 `{ width: "120pt", height: "60pt" }` |
| `title` | 字符串 | 图片标题 |
| `fill_color` | VML 颜色，例如 `"#00FF00"` | 图片背后的填充 |
| `opacity` | 字符串，例如 `"50%"` | 填充的不透明度 |
| `wrap` | `true` | 关闭图片周围的文字环绕 |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**定位图片**：`p.image(media, pH:, pV:, wrap: true, width:, height:)` 以 EMU 为单位，把图片锚定在相对于栏（`pH`）和段落（`pV`）的偏移位置。`pH`、`pV` 和 `wrap` 必须同时提供，否则图片会以内联方式放置。

**页眉图片**：`h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`。`wrap: "none"` 关闭环绕，其他任何值都会让文字以四周型环绕图片。

**0.1.0 中的已知问题**：定位图片、页眉图片和表格单元格图片的宽度和高度都按 `width * height * 14.92` EMU 计算。它们总是正方形，而且通常很小。用 `style:` 设置尺寸的内联图片不受影响。变通方法参见[最佳实践](Best-Practices#images)。

## <a id="headers-and-footers"></a>页眉和页脚

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- 页眉可以包含 `text`、`br`、`table`、`image` 和 `page_numbers`。页脚可以包含 `text`、`br`、`table` 和 `page_numbers`。
- 传给 `Header.new` 或 `Footer.new` 的选项作用于其中的文本。
- `add_header` 和 `add_footer` 作用于第 2 页及之后的页面。第 1 页始终使用 `add_first_header` 和 `add_first_footer`。参见[配置](Configuration#first-page)。

## <a id="page-numbers"></a>页码

`page_numbers(name: "Page", align: "center")` 先写入 `name` 文本，再写入一个 PAGE 域。页码使用 Arial 10 磅、深灰色。`align` 默认为 `"right"`。

已知问题：域代码中 `\* MERGEFORMAT` 缺少反斜杠。

## <a id="table-of-contents"></a>目录

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`。
- 提供页码时，该行会在 `pos` 处得到一个右对齐制表位（单位为二十分之一磅，9360 即 6.5 英寸），并用 `leader` 字符填充。前导符的取值来自 WordprocessingML，例如 `"dot"`、`"hyphen"`、`"underscore"`、`"middleDot"` 和 `"none"`。
- 不提供页码时，该行是制表符后面的文本。
- 页码就是你传入的值。docxi 不会创建 Word 的 TOC 域，也不会计算页码。

## <a id="hyperlinks"></a>超链接

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` 返回一个关系 ID，`hyperlink(text, url, id)` 负责放置链接。链接目标来自关系，所以要传入为目标 URL 返回的那个 ID。只支持外部链接。

## <a id="frames-text-boxes"></a>框架（文本框）

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- 框架是 VML 文本框，带有固定的浅灰色半透明填充和黑色轮廓。文字在其上方和下方排布。
- 在框架中可以调用 `text`、`br`、`tab`、`image` 和 `frame`。
- `spacing: true` 在文本前添加间距。

## <a id="output"></a>输出

`word.render` 返回一个包含完整 `.docx` 包的 `StringIO`。压缩包中不含 Zip64 记录，因此 Google Docs 和较旧的 LibreOffice 版本可以打开它。[架构](Architecture)列出了包中的每个部件。

## <a id="related-pages"></a>相关页面

- [使用示例](Usage-Examples)
- [API 参考](API-Reference)
- [路线图](Roadmap)：计划在 0.2.0 中修复的已知问题

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
