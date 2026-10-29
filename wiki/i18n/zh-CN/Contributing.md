<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# 为 docxi 做贡献

欢迎提交 bug 报告、修复、文档和翻译。本页说明如何报告问题，以及 pull request 需要满足哪些要求。要搭建本地开发副本，参见[开发环境搭建](Development-Setup)。

## <a id="reporting-a-bug"></a>报告 bug

在 [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) 提交 issue，并附上：

- 一个能重现问题的简短 Ruby 脚本
- Ruby、docxi、nokogiri 和 rubyzip 的版本（如果使用 Rails，也请注明其版本）
- 你预期的结果，以及实际得到的结果。如果文档损坏，请说明你是用哪个程序打开的

安全问题请私下报告。参见[安全](Security)。

## <a id="pull-requests"></a>Pull request

1. Fork 仓库并创建分支。
2. 进行修改，并与周围代码的风格保持一致。
3. 添加或更新测试，然后运行 `bundle exec rake test`。
4. 如果生成的 XML 是有意改变的，请运行 `bundle exec rake golden`，并在 pull request 中说明这一变化。基准文件保护着现有用户所依赖的输出。
5. 如果行为发生变化，请更新 `CHANGELOG.md` 和文档。
6. 提交 pull request。CI 会在 Ruby 2.5 到 4.0 和 Rails 5.2 到 8.1 上运行测试，并检查 LibreOffice 能否打开示例文档。

会改变现有文档的修改（例如计划在 0.2.0 中进行的修复）应放在单独的 pull request 中，这样基准文件的差异更容易审查。

## <a id="documentation-and-translations"></a>文档和翻译

- 英文 README 是 `docs/i18n/` 中 15 个译本的来源。修改它时，请遵循仓库中的 `docs/i18n/TRANSLATION_GUIDE.md`。
- 本 wiki 在仓库的 `wiki/` 文件夹中编写，并从那里发布。`wiki/README.md` 说明了具体做法。

## <a id="related-pages"></a>相关页面

- [开发环境搭建](Development-Setup)
- [架构](Architecture)
- [路线图](Roadmap)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
