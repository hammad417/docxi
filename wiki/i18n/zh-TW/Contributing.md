<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# 參與 docxi 的貢獻

歡迎提交錯誤回報、修正、文件和翻譯。本頁說明如何回報問題，以及 pull request 需要符合哪些要求。若要設定本機開發環境，請參閱[開發環境設定](Development-Setup)。

## <a id="reporting-a-bug"></a>回報錯誤

在 [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) 提交 issue，並附上：

- 一個能重現問題的簡短 Ruby 指令碼
- Ruby、docxi、nokogiri 和 rubyzip 的版本（如果使用 Rails，也請註明其版本）
- 你預期的結果，以及實際得到的結果。如果文件損壞，請說明你是用哪個程式開啟的

安全性問題請私下回報。請參閱[安全性](Security)。

## <a id="pull-requests"></a>Pull request

1. Fork 儲存庫並建立分支。
2. 進行修改，並與周圍程式碼的風格保持一致。
3. 新增或更新測試，然後執行 `bundle exec rake test`。
4. 如果產生的 XML 是刻意變更的，請執行 `bundle exec rake golden`，並在 pull request 中說明這項變更。基準檔案保護著現有使用者所依賴的輸出。
5. 如果行為有所變更，請更新 `CHANGELOG.md` 和文件。
6. 提交 pull request。CI 會在 Ruby 2.5 到 4.0 和 Rails 5.2 到 8.1 上執行測試，並確認 LibreOffice 能開啟範例文件。

會改變現有文件的修改（例如預計在 0.2.0 進行的修正）應放在獨立的 pull request 中，這樣基準檔案的差異比較容易審查。

## <a id="documentation-and-translations"></a>文件和翻譯

- 英文 README 是 `docs/i18n/` 中 15 個譯本的來源。修改它時，請遵循儲存庫中的 `docs/i18n/TRANSLATION_GUIDE.md`。
- 本 wiki 在儲存庫的 `wiki/` 資料夾中撰寫，並從那裡發布。`wiki/README.md` 說明了具體做法。

## <a id="related-pages"></a>相關頁面

- [開發環境設定](Development-Setup)
- [架構](Architecture)
- [路線圖](Roadmap)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
