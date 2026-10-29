<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 路線圖與已知問題

docxi 沒有發布時程，也沒有承諾的路線圖。目前有記錄的規劃工作只有修正下列已知問題。它們預計在 0.2.0 修正，沒有確定的日期。之所以沒有放進 0.1.0，是因為修正它們會改變現有使用者已經在產生的文件。

## <a id="known-issues-planned-for-020"></a>預計在 0.2.0 修正的已知問題

| 問題 | 0.1.0 的因應措施 |
|---|---|
| 定位圖片、頁首圖片和表格儲存格圖片的兩個尺寸都以 `width * height * 14.92` EMU 計算 | 使用以 `style:` 設定尺寸的內嵌圖片，或參閱[最佳做法](Best-Practices#images) |
| 表格層級的 `borders:` 會套用到每個儲存格，而不是表格的外框 | 如果只需要外框，只在邊緣儲存格的外側設定 `borders:` |
| 頁碼功能變數的代碼中 `\* MERGEFORMAT` 缺少反斜線 | 無 |
| 表格儲存格中的頁碼前面多出「GlobalOptions」 | 將頁碼放在頁首或頁尾中 |
| 圖片一律以 PNG 組件儲存，而且 `[Content_Types].xml` 只宣告前兩個頁首和頁尾 | 使用 PNG 圖片，並且最多使用兩個頁首和兩個頁尾 |
| 只傳入部分邊界時，其餘的會寫成空值 | 同時傳入 `top`、`right`、`bottom` 和 `left` |
| `docProps/core.xml` 記錄的是本地時間卻帶有 UTC 標記，而且從不寫入作者 | 無 |
| `Docxi.test` 會引發 `NoMethodError` | 不要呼叫它，它是遺留的輔助方法 |
| 沒有 `page: true` 的 `document.br` 和 `document.tab` 會產生無效的 XML | 在段落中使用 `p.br` 和 `p.tab` |

## <a id="not-planned"></a>未列入規劃

這些不在上面的清單中，目前也還沒有做出任何決定：

- 編輯現有文件或填入範本
- 編號清單、標題樣式或自動目錄
- 其他頁面大小或橫向頁面

## <a id="related-pages"></a>相關頁面

- [變更紀錄](Changelog)
- [參與貢獻](Contributing)
- [常見問題](FAQ)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
