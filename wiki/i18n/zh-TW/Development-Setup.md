<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# 開發環境設定：執行 docxi 的測試和工具

本頁協助你執行 docxi 的本機副本，並介紹測試、範例文件和 CI 設定。

## <a id="get-the-code"></a>取得程式碼

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

開發時可以使用 Ruby 2.5 以上的任何版本。開發用的相依套件是 minitest 和 rake。

## <a id="rake-tasks"></a>Rake 工作

| 命令 | 作用 |
|---|---|
| `bundle exec rake test` | 執行測試套件（預設工作） |
| `bundle exec rake samples` | 將範例文件寫入 `tmp/samples/*.docx`，以便在 Word、LibreOffice 或 Google Docs 中開啟 |
| `bundle exec rake golden` | 根據目前的程式碼重寫 `test/golden`。只在刻意變更輸出後使用 |
| `bundle exec rake build` | 將 gem 建置到 `pkg/` 中 |

## <a id="running-the-tests-with-rails-loaded"></a>在載入 Rails 的情況下執行測試

`gemfiles/` 資料夾中每個 Rails 版本都有一個 Gemfile，另外還有幾個用於較舊 rubyzip 和 nokogiri 版本的 Gemfile：

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` 會在測試前載入整個框架，也會將 docxi 的 `blank?` 規則與 ActiveSupport 的規則進行比對。

## <a id="what-the-tests-cover"></a>測試涵蓋的內容

| 檔案 | 檢查內容 |
|---|---|
| `test/test_golden_output.rb` | 兩份範例文件的每個組件都與 0.0.6 產生的相同 |
| `test/test_package.rb` | XML 格式正確、關聯指向存在的組件、每個組件都有內容類型，而且沒有 Zip64 記錄 |
| `test/test_media.rb` | 來自路徑、IO 物件和 URL 的圖片，以及 `\|` 字串絕不會被當作命令執行 |
| `test/test_fixed_features.rb` | `tab`、`frame` 和頁首中的頁碼 |
| `test/test_util.rb` | `Docxi::Util.blank?` 遵循 ActiveSupport 的規則 |
| `test/test_rails_independence.rb` | docxi 可以在沒有 ActiveSupport 的情況下執行 |
| `test/test_warnings.rb` | 渲染時 `lib/` 不會發出淘汰警告 |

範例文件位於 `test/support/sample_documents.rb`，`test/support/package.rb` 會直接從 zip 結構中讀取 `.docx`。

## <a id="continuous-integration"></a>持續整合

`.github/workflows/ci.yml` 會在每次推送到 `master` 以及每個 pull request 上執行，涵蓋：

- 不載入 Rails 的 Ruby 2.5、2.6 和 4.0
- Rails 5.2（Ruby 2.5）、6.0（2.7）、6.1（3.0）、7.0（3.1）、7.1（3.2）、7.2（3.3）、8.0（3.4）和 8.1（4.0）
- Rails `main`，允許失敗
- rubyzip 1.3 和 2.4，以及 nokogiri 1.11
- 一個以 LibreOffice 將範例文件轉換成 PDF 的工作

版本發布的說明請參閱[部署](Deployment)。

## <a id="related-pages"></a>相關頁面

- [參與貢獻](Contributing)
- [架構](Architecture)
- [部署](Deployment)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
