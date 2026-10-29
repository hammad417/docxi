<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 最佳实践：避开常见的意外

这些习惯直接源自 docxi 的工作方式，其中大多数可以避免[故障排除](Troubleshooting)中列出的问题。它们按主题分组。页面设置的建议适用于所有人，安全和性能方面的建议对生产环境中的应用最为重要。

## <a id="page-setup"></a>页面设置

- **同时传入全部四个页边距**。传给 `Docxi::Document.new` 的哈希会替换默认值，`top`、`right`、`bottom` 和 `left` 中漏掉的任何一项都会写成空值。
- **设置首页的页眉和页脚**。docxi 始终开启独立的首页。如果第 1 页需要页眉或页脚，还要调用 `add_first_header` 和 `add_first_footer`。

## <a id="text"></a>文本

- **开头或结尾有空格时使用 `space: true`**。否则 Word 会丢掉这些空格，两个文本片段的词就会连在一起。
- **把共用的格式放在段落上**。传给 `p` 的选项（例如 `size:` 或 `font:`）会作用于其中的每个文本片段，需要重复的内容就更少。
- **换行用 `p.br`，分页用 `document.br(page: true)`**。不带 `page: true` 的 `document.br` 会生成无效的 XML。

## <a id="tables"></a>表格

- **在行或单元格上设置边框，而不是在表格上**。在 0.1.0 中，表格级的 `borders:` 会应用到每个单元格，而不是表格的外框。
- **为每一列提供 `columns_width`**。单元格从中获取宽度，表格就能保持你预期的布局。

## <a id="images"></a>图片

- **使用 PNG 文件**。每张图片都以 PNG 部件存储。
- **用 `style:` 设置内联图片的尺寸**，例如 `style: { width: "120pt", height: "60pt" }`。在 0.1.0 中，这是唯一能按预期工作的图片尺寸设置方式。
- **如果文档可能渲染多次，请传入路径，而不是已打开的 `File`**。IO 图片在 `render` 期间读取，第二次渲染会得到空图片。
- **0.2.0 之前的页眉图片、表格单元格图片和定位图片**。它们的宽度和高度都按 `width * height * 14.92` EMU 计算，所以结果是正方形。要得到边长为 `S` 英寸的正方形图片，请把 `width` 和 `height` 都设为 `Math.sqrt(S * 914400 / 14.92).round`，一英寸约为 248。

## <a id="security"></a>安全

- **不要把不可信的字符串传给 `add_media`**。传入路径会让 docxi 读取服务器上的该文件，传入 URL 会让你的服务器下载它。请传入你自己打开的 IO，或者先检查该值。参见[安全](Security)。
- **文本会自动转义**。传给 `text` 的字符串会作为 XML 文本写入，因此用户输入无法改变文档结构。

## <a id="performance-and-reliability"></a>性能与可靠性

- **在后台任务中构建大型文档**。调用 `render` 时，整个文档都在内存中构建。参见[集成](Integrations)。
- **只渲染一次，然后复用字节内容**。如果需要多次发送或存储文档，请保留 `word.render.string`。
- **用 `gem "docxi", "~> 0.1"` 锁定兼容的版本**。
- **测试你生成的内容**。`.docx` 是由 XML 文件组成的 zip 压缩包，所以测试可以解压它并检查 XML。docxi 自己的测试就是这样做的。参见[架构](Architecture)。

## <a id="related-pages"></a>相关页面

- [功能](Features)
- [配置](Configuration)
- [故障排除](Troubleshooting)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
