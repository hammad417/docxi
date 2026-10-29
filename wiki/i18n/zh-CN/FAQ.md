<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 常见问题：在 Ruby 和 Rails 中生成 Word 文档

这里是对最常见问题的简短回答，每个回答都附有详细说明的链接。

**需要安装 Microsoft Word 或 LibreOffice 吗？**
不需要。docxi 自己写出 XML 和 zip 压缩包。

**docxi 依赖 Rails 吗？**
不依赖。它可以在任何 Ruby 程序中使用。测试也会在加载 Rails 5.2 到 8.1 的情况下运行。参见[集成](Integrations)。

**如何在 Ruby on Rails 中生成 Word 文档？**
在控制器或后台任务中构建文档，然后用 `send_data` 发送 `word.render.string`。参见[集成](Integrations)。

**支持哪些 Ruby 和 Rails 版本？**
Ruby 2.5 及更高版本。Rails 是可选的，CI 测试了 Rails 5.2 到 8.1。参见[安装](Installation)。

**docxi 能编辑已有的 .docx 或填充 Word 模板吗？**
不能，它只能创建新文档。如果需要带合并域的模板，请使用基于模板的 gem，例如 sablon。

**目录是根据标题自动生成的吗？**
不是。每一行及其页码都由你添加。参见[功能](Features#table-of-contents)。

**可以使用编号列表或标题样式吗？**
0.1.0 中不行。段落可以设为项目符号列表项（`ul: true`），标题可以用 `size:` 和 `bold:` 等文本选项设置格式。

**可以更改页面大小或方向吗？**
暂时不行。页面为纵向 US Letter，只能修改页边距。参见[配置](Configuration)。

**支持哪些图片格式？**
PNG。每张图片都以 PNG 部件存储。

**可以从 URL 添加图片吗？**
可以。`add_media` 接受 `http`、`https` 和 `ftp` URL，并在调用 `render` 时下载。不要未经检查就传入来自用户的 URL。参见[安全](Security)。

**可以把文档直接写入文件吗？**
`render` 返回一个 `StringIO`。用 `File.binwrite("out.docx", word.render.string)` 写入它的 `.string`。

**docxi 能以流式方式输出超大文档吗？**
不能。调用 `render` 时，整个文档都在内存中构建。参见[部署](Deployment)。

**可以设置从右到左的段落吗？**
0.1.0 中没有从右到左的选项，文本按你传入的原样写入。

**已知问题列在哪里？**
在[路线图](Roadmap)上，以及 README 的“Limitations and known issues”（限制和已知问题）一节中。

## <a id="related-pages"></a>相关页面

- [故障排除](Troubleshooting)
- [功能](Features)
- [快速入门](Getting-Started)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
