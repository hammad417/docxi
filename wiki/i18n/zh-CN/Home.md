<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi wiki：用 Ruby 生成 Word（.docx）文档

docxi 是一个 Ruby gem，用代码生成 Microsoft Word `.docx` 文件。它支持段落、带格式的文本、表格、图片、页眉和页脚、页码、超链接以及简单的目录。它不需要 Microsoft Word、LibreOffice 或模板，在普通 Ruby 程序和 Ruby on Rails 应用中都能使用。

[README](https://github.com/hammad417/docxi#readme) 是简短的介绍。本 wiki 涵盖所有选项，提供可直接复制的示例，说明 docxi 的内部原理，并在出现问题时提供帮助。

当前版本：**0.1.1**，适用于 Ruby 2.5 及更高版本。参见[更新日志](Changelog)。

## <a id="start-here"></a>从这里开始

**初次使用 docxi**

- [快速入门](Getting-Started)：安装 gem 并创建第一个文档
- [安装](Installation)：环境要求、Bundler，以及从 0.0.x 升级
- [使用示例](Usage-Examples)：报告、发票、信函和由数据生成的表格，可直接复制

**构建文档**

- [功能](Features)：每种元素及其支持的选项
- [配置](Configuration)：页边距、页面设置和单位
- [API 参考](API-Reference)：类和方法一览
- [最佳实践](Best-Practices)：避开常见意外的习惯

**出现问题时**

- [故障排除](Troubleshooting)：错误信息，以及显示不正常的文档
- [常见问题](FAQ)：常见问题的简短回答
- [安全](Security)：报告问题以及处理不可信输入

**深入了解**

- [架构](Architecture)：Ruby 调用如何变成 `.docx` 包
- [集成](Integrations)：Rails 控制器、Rack 应用和后台任务
- [部署](Deployment)：在生产环境中运行 docxi，以及版本如何发布
- [参与贡献](Contributing)和[开发环境搭建](Development-Setup)
- [路线图](Roadmap)、[更新日志](Changelog)和[致谢](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi 能做什么，不能做什么

docxi 用于创建新文档。它不能打开或编辑已有的 `.docx`，不能填充 Word 模板，也不能执行邮件合并。它没有编号列表、标题样式或自动目录，页面始终是纵向的 US Letter。计划在 0.2.0 中修复的已知问题列在[路线图](Roadmap)上。

## <a id="credits"></a>致谢

docxi 由 [Phaedra Solutions](https://www.phaedrasolutions.com/) 的 [Hammad Maqbool](https://github.com/hammad417) 维护。它于 2016 年在 Timur Gabdrakipov 的 [docxer](https://github.com/gtimti/docxer) 基础上起步。详情参见[致谢](Credits)。

## <a id="languages"></a>语言

英文版是本 wiki 的参考版本。各语言译本的链接位于侧边栏和每个页面的顶部。

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [报告问题](https://github.com/hammad417/docxi/issues)
