<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# 开发环境搭建：运行 docxi 的测试和工具

本页帮助你运行 docxi 的本地副本，并介绍测试、示例文档和 CI 配置。

## <a id="get-the-code"></a>获取代码

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

开发时可以使用 Ruby 2.5 及以上的任何版本。开发依赖是 minitest 和 rake。

## <a id="rake-tasks"></a>Rake 任务

| 命令 | 作用 |
|---|---|
| `bundle exec rake test` | 运行测试套件（默认任务） |
| `bundle exec rake samples` | 把示例文档写入 `tmp/samples/*.docx`，以便在 Word、LibreOffice 或 Google Docs 中打开 |
| `bundle exec rake golden` | 根据当前代码重写 `test/golden`。仅在有意改变输出后使用 |
| `bundle exec rake build` | 把 gem 构建到 `pkg/` 中 |

## <a id="running-the-tests-with-rails-loaded"></a>在加载 Rails 的情况下运行测试

`gemfiles/` 文件夹中每个 Rails 版本都有一个 Gemfile，另外还有几个用于较旧 rubyzip 和 nokogiri 版本的 Gemfile：

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` 会在测试前加载整个框架，还会把 docxi 的 `blank?` 规则与 ActiveSupport 的进行比较。

## <a id="what-the-tests-cover"></a>测试覆盖的内容

| 文件 | 检查内容 |
|---|---|
| `test/test_golden_output.rb` | 两个示例文档的每个部件都与 0.0.6 生成的一致 |
| `test/test_package.rb` | XML 格式良好，关系指向存在的部件，每个部件都有内容类型，并且没有 Zip64 记录 |
| `test/test_media.rb` | 来自路径、IO 对象和 URL 的图片，以及 `\|` 字符串永远不会作为命令执行 |
| `test/test_fixed_features.rb` | `tab`、`frame` 和页眉中的页码 |
| `test/test_util.rb` | `Docxi::Util.blank?` 遵循 ActiveSupport 的规则 |
| `test/test_rails_independence.rb` | docxi 可以在没有 ActiveSupport 的情况下运行 |
| `test/test_warnings.rb` | 渲染时 `lib/` 不会发出弃用警告 |

示例文档位于 `test/support/sample_documents.rb`，`test/support/package.rb` 直接从 zip 结构中读取 `.docx`。

## <a id="continuous-integration"></a>持续集成

`.github/workflows/ci.yml` 会在每次推送到 `master` 以及每个 pull request 上运行，涵盖：

- 不加载 Rails 的 Ruby 2.5、2.6 和 4.0
- Rails 5.2（Ruby 2.5）、6.0（2.7）、6.1（3.0）、7.0（3.1）、7.1（3.2）、7.2（3.3）、8.0（3.4）和 8.1（4.0）
- Rails `main`，允许失败
- rubyzip 1.3 和 2.4，以及 nokogiri 1.11
- 一个用 LibreOffice 把示例文档转换为 PDF 的任务

版本发布在[部署](Deployment)中介绍。

## <a id="related-pages"></a>相关页面

- [参与贡献](Contributing)
- [架构](Architecture)
- [部署](Deployment)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
