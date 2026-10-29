<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# 配置 docxi：页边距、页面设置和单位

docxi 没有配置文件、环境变量或全局设置，所有设置都按文档或按元素指定。本页介绍文档级的设置以及各选项使用的单位。每种元素的选项列在[功能](Features)页面。

## <a id="page-margins"></a>页边距

页边距在创建文档时传入：

```ruby
require "docxi"

# 所有数值的单位都是二十分之一磅（1440 = 1 英寸）。
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| 选项 | 默认值 | 含义 |
|---|---|---|
| `top` | `720`（0.5 英寸） | 上边距 |
| `right` | `1440`（1 英寸） | 右边距 |
| `bottom` | `720`（0.5 英寸） | 下边距 |
| `left` | `1440`（1 英寸） | 左边距 |
| `header` | `190` | 页面上边缘到页眉的距离 |
| `footer` | `200` | 页面下边缘到页脚的距离 |

你传入的哈希会整体替换默认值，所以请同时设置 `top`、`right`、`bottom` 和 `left`。这四项中漏掉的任何一项都会写成空值。`header` 和 `footer` 缺失或为空时，会回退为 190 和 200。

## <a id="page-size-and-orientation"></a>页面大小和方向

页面为纵向 US Letter：12240 × 15840 个二十分之一磅，即 8.5 × 11 英寸。在 0.1.0 中无法更改页面大小和方向。

## <a id="first-page"></a>首页

docxi 始终开启 Word 的“首页不同”设置。用 `add_header` 和 `add_footer` 添加的页眉和页脚从第 2 页开始使用。第 1 页使用 `add_first_header` 和 `add_first_footer`，因此如果首页需要页眉或页脚，也要设置它们。参见[功能](Features#headers-and-footers)。

## <a id="what-every-document-contains"></a>每个文档都包含的内容

- 默认字体为 Arial 12 磅。
- 定义的样式：`Normal`、`DefaultParagraphFont`、`TableNormal`、`TableCustom` 和 `NoList`。
- 项目符号列表项（`ul: true`）使用短横线作为项目符号。
- 文档属性把应用程序名称记录为“Docxi Ruby Gem”，并记录文档渲染的时间。

## <a id="units"></a>单位

| 选项 | 单位 |
|---|---|
| 文本 `size:` | 磅（docxi 会将其乘以 2，换算为 Word 的半磅） |
| 页边距、段落 `left:` 和 `right:`、表格 `iwidth:`、行 `height:` | 二十分之一磅（每英寸 1440） |
| 单元格 `borders:` 的尺寸 | 八分之一磅 |
| `columns_width:` 和单元格 `width:` | 乘以 14.92 得到二十分之一磅，在 96 dpi 下接近像素值 |
| 内联图片 `style:` | CSS 长度，例如 `"120pt"` |
| 定位图片 `pH:` 和 `pV:` | EMU（每英寸 914400） |
| 文本和填充颜色 | 不带 `#` 的十六进制值，例如 `"8B0000"` |
| 图片 `fill_color:` | 带 `#` 的 VML 颜色，例如 `"#00FF00"` |

## <a id="related-pages"></a>相关页面

- [功能](Features)
- [API 参考](API-Reference)
- [最佳实践](Best-Practices)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
