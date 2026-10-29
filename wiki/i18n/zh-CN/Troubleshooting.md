<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 故障排除：安装错误、运行时错误和损坏的输出

在下面找到你遇到的症状。旧应用中的大多数问题都来自 docxi 0.0.x，升级到 0.1.0 即可解决，无需修改代码。

## <a id="installation"></a>安装

**在 Rails 6+ 应用中，Bundler 找不到兼容的 `nokogiri` 版本**
docxi 0.0.6 锁定了 `nokogiri < 1.7`，而 Rails 6.0 及之后的每个版本都需要 nokogiri 1.8.5 或更高版本。请使用 `gem "docxi", "~> 0.1"`。

**Bundler 找不到兼容的 `rubyzip` 版本**
docxi 0.0.6 锁定了 `rubyzip ~> 0.9.9`，与任何需要更新版 rubyzip 的 gem 冲突。docxi 0.1.0 接受 rubyzip 1.3 到 3.x。

**Bundler 一直安装 0.0.6**
原因可能是 Gemfile 把 docxi 固定在 0.0.x（例如 `"~> 0.0.6"`），也可能是 Ruby 低于 0.1.0 要求的 2.5。请把 Gemfile 中的这一行改为 `gem "docxi", "~> 0.1"`，并检查 `ruby -v`。

**nokogiri 编译失败**
没有与你的平台匹配的预编译 nokogiri gem，所以它正在从源码构建。请安装 C 编译器，并按照 [nokogiri 安装指南](https://nokogiri.org/tutorials/installing_nokogiri.html)操作。

## <a id="errors-while-building-a-document"></a>构建文档时出错

**在 Rails 之外出现 `NoMethodError: undefined method 'present?'`**
docxi 0.0.6 假定 ActiveSupport 已经加载。0.1.0 不需要它。

**调用 `tab`、`frame` 或页眉的 `page_numbers` 时出现 `NameError`**
这些功能在 0.0.6 中无法使用，在 0.1.0 中可以正常工作，包括 `document.frame("text")`。

**添加图片时出现 `Errno::ENOENT`**
该路径不存在（相对于程序运行时所在的目录）。在 Ruby 3 上使用 docxi 0.0.6 时，图片 URL 也会以同样的方式失败，因为 0.0.6 使用的是 `Kernel#open`。0.1.0 用 `URI#open` 下载 URL。

## <a id="documents-that-dont-look-right"></a>文档显示不正常

**Google Docs 或 LibreOffice 无法打开文件**
docxi 0.0.6 配合 rubyzip 3 时会在每个条目中写入 Zip64 记录，而这些程序会拒绝这种记录。0.1.0 写出的压缩包不含这些记录。

**首页没有页眉或页脚**
docxi 为首页单独设置页眉和页脚。请用 `add_first_header` 和 `add_first_footer` 设置它们。参见[配置](Configuration#first-page)。

**缺少某个页边距，或页面布局看起来不对**
设置页边距时，请同时传入 `top`、`right`、`bottom` 和 `left`。漏掉的那些会被写成空值。

**文本片段之间的空格消失了**
Word 会丢掉文本片段开头或结尾的空格。请在该 `text` 调用中添加 `space: true`。

**页眉中、表格单元格中或使用 `pH`/`pV` 的图片很小或者是正方形**
0.1.0 中的已知问题：它们的两个尺寸都按 `width * height * 14.92` EMU 计算。请使用以 `style:` 设置尺寸的内联图片，或参见[最佳实践](Best-Practices#images)中的变通方法。

**同一文档第二次渲染时缺少图片**
从已打开的 `File` 或其他 IO 添加的图片在 `render` 期间读取，第二次渲染读到的是空的流。请改为传入路径，或者只渲染一次。

**在 `table` 上设置的边框出现在每个单元格周围**
已知问题：表格级的 `borders:` 会应用到每个单元格，而不是表格的外框。如果只想绘制外框，请只在边缘单元格上为其外侧设置 `borders:`。

**表格单元格中的页码以“GlobalOptions”开头**
这是 `TableCell#page_numbers` 的已知问题。请改为把页码放在页脚或页眉中。

**JPEG 或 GIF 图片显示不正确或不显示**
无论实际格式是什么，每张图片都以 PNG 部件和 PNG 内容类型存储，因此某些程序可能无法显示其他格式。请在添加之前把图片转换为 PNG。

## <a id="inspecting-a-generated-file"></a>检查生成的文件

`.docx` 是一个 zip 压缩包，所以可以查看其内部：

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

如果问题仍然存在，请[提交 issue](https://github.com/hammad417/docxi/issues)，附上能重现问题的简短脚本，以及 Ruby、docxi、nokogiri 和 rubyzip 的版本。

## <a id="related-pages"></a>相关页面

- [常见问题](FAQ)
- [最佳实践](Best-Practices)
- [安装](Installation)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
