<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 路线图和已知问题

docxi 没有发布计划，也没有承诺的路线图。目前有记录的计划工作只有修复下面列出的已知问题。它们计划在 0.2.0 中修复，没有确定日期。之所以没有放进 0.1.0，是因为修复它们会改变现有用户已经在生成的文档。

## <a id="known-issues-planned-for-020"></a>计划在 0.2.0 中修复的已知问题

| 问题 | 0.1.0 中的变通方法 |
|---|---|
| 定位图片、页眉图片和表格单元格图片的两个尺寸都按 `width * height * 14.92` EMU 计算 | 使用以 `style:` 设置尺寸的内联图片，或参见[最佳实践](Best-Practices#images) |
| 表格级的 `borders:` 会应用到每个单元格，而不是表格的外框 | 如果只需要外框，只在边缘单元格的外侧设置 `borders:` |
| 页码域的域代码中 `\* MERGEFORMAT` 缺少反斜杠 | 无 |
| 表格单元格中的页码前面带有“GlobalOptions” | 把页码放在页眉或页脚中 |
| 图片始终以 PNG 部件存储，并且 `[Content_Types].xml` 只声明前两个页眉和页脚 | 使用 PNG 图片，并且最多使用两个页眉和两个页脚 |
| 只传入部分页边距时，其余的会写成空值 | 同时传入 `top`、`right`、`bottom` 和 `left` |
| `docProps/core.xml` 记录的是本地时间却带有 UTC 标记，并且从不写入作者 | 无 |
| `Docxi.test` 会抛出 `NoMethodError` | 不要调用它，它是遗留的辅助方法 |
| 不带 `page: true` 的 `document.br` 和 `document.tab` 会生成无效的 XML | 在段落中使用 `p.br` 和 `p.tab` |

## <a id="not-planned"></a>未列入计划

这些不在上面的列表中，也还没有做出任何决定：

- 编辑已有文档或填充模板
- 编号列表、标题样式或自动目录
- 其他页面大小或横向页面

## <a id="related-pages"></a>相关页面

- [更新日志](Changelog)
- [参与贡献](Contributing)
- [常见问题](FAQ)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
