<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 常見問題：在 Ruby 和 Rails 中產生 Word 文件

這裡是最常見問題的簡短解答，每個解答都附有詳細說明的連結。

**需要安裝 Microsoft Word 或 LibreOffice 嗎？**
不需要。docxi 會自行寫出 XML 和 zip 壓縮檔。

**docxi 需要 Rails 嗎？**
不需要。它可以在任何 Ruby 程式中使用。測試也會在載入 Rails 5.2 到 8.1 的情況下執行。請參閱[整合](Integrations)。

**如何在 Ruby on Rails 中產生 Word 文件？**
在控制器或背景工作中建立文件，然後用 `send_data` 傳送 `word.render.string`。請參閱[整合](Integrations)。

**支援哪些 Ruby 和 Rails 版本？**
Ruby 2.5 以上版本。Rails 是選用的，CI 測試了 Rails 5.2 到 8.1。請參閱[安裝](Installation)。

**docxi 能編輯現有的 .docx 或填入 Word 範本嗎？**
不能，它只能建立新文件。如果需要含合併欄位的範本，請使用以範本為基礎的 gem，例如 sablon。

**目錄是根據標題自動產生的嗎？**
不是。每個項目及其頁碼都由你加入。請參閱[功能](Features#table-of-contents)。

**可以使用編號清單或標題樣式嗎？**
0.1.0 不行。段落可以設為項目符號清單項目（`ul: true`），標題可以用 `size:` 和 `bold:` 等文字選項設定格式。

**可以變更頁面大小或方向嗎？**
目前還不行。頁面為直向的 US Letter，只能變更邊界。請參閱[設定](Configuration)。

**支援哪些圖片格式？**
PNG。每張圖片都以 PNG 組件儲存。

**可以從 URL 加入圖片嗎？**
可以。`add_media` 接受 `http`、`https` 和 `ftp` URL，並在呼叫 `render` 時下載。不要未經檢查就傳入來自使用者的 URL。請參閱[安全性](Security)。

**可以直接將文件寫入檔案嗎？**
`render` 傳回一個 `StringIO`。請用 `File.binwrite("out.docx", word.render.string)` 寫入它的 `.string`。

**docxi 能以串流方式輸出非常大的文件嗎？**
不能。呼叫 `render` 時，整份文件都會在記憶體中建立。請參閱[部署](Deployment)。

**可以設定從右到左的段落嗎？**
0.1.0 沒有從右到左的選項，文字會依你傳入的原樣寫入。

**已知問題列在哪裡？**
在[路線圖](Roadmap)上，以及 README 的「Limitations and known issues」（限制與已知問題）一節中。

## <a id="related-pages"></a>相關頁面

- [疑難排解](Troubleshooting)
- [功能](Features)
- [快速入門](Getting-Started)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
