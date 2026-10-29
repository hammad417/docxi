<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 更新日志和版本历史

本页概述每个版本。完整的说明位于仓库中的 [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md)，两者如有出入，以它为准。

## <a id="releases"></a>版本

| 版本 | 日期 | 说明 |
|---|---|---|
| 0.1.1 | 2026-10-29 | gem 的摘要和描述更准确，并新增文档链接。代码没有变化 |
| 0.1.0 | 2026-10-28 | 兼容当前的 Ruby、Rails 和 rubyzip。生成的文档与 0.0.6 相同 |
| 0.1.0.rc1 | 2026-10-28 | 0.1.0 的预发布版本 |
| 0.0.6 | 2016-10-18 | 0.1.0 更新之前的最后一个版本 |
| 0.0.2 | 2016-10-14 | 没有记录发布说明 |
| 0.0.1 | 2016-10-14 | 第一个版本。没有记录发布说明 |

0.0.3 到 0.0.5 从未发布到 RubyGems。

## <a id="011"></a>0.1.1

代码没有变化，生成的文档与 0.1.0 相同。

- gem 的摘要和描述现在说明了 docxi 的实际用途。之前它们把 docxi 称为模板生成器。
- RubyGems 把本 wiki 链接为 gem 的文档。

## <a id="010"></a>0.1.0

**兼容性**

- 需要 Ruby 2.5 或更高版本。使用更旧 Ruby 的应用会停留在 0.0.6。
- 有没有 Rails 都可以使用，并与 Rails 5.2 到 8.1 一起测试过。
- nokogiri 的要求从 `>= 1.6.3, < 1.7` 改为 `>= 1.8, < 2`，前者与 Rails 6.0 以来的每个版本都冲突。
- rubyzip 的要求从 `~> 0.9.9` 改为 `>= 1.3, < 4`。
- 不再调用 ActiveSupport 的 `blank?` 和 `present?`，这两个方法只有在 Rails 加载它们之后才可用。

**修复**

- 使用 rubyzip 3 写出的文档不再包含 Zip64 记录，Google Docs 和较旧的 LibreOffice 版本会拒绝这种记录。
- 在 Ruby 3 上，使用图片 URL 调用 `add_media` 恢复正常。
- `add_media` 不再把以 `|` 开头的字符串当作 shell 命令执行。
- `tab` 和 `frame` 不再抛出 `NameError`。
- `document.frame("text")` 会生成一个包含该文本的文本框。
- 页眉中的 `page_numbers` 不再抛出 `NameError`。

**gem 包**

- 新的维护者，以及新的主页和更新日志链接。
- 发布版本需要 MFA。

## <a id="related-pages"></a>相关页面

- [路线图](Roadmap)
- [安装](Installation#upgrading-from-00x)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
