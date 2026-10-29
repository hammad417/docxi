<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 最佳做法：避開常見的意外

這些習慣直接來自 docxi 的運作方式，其中大多數可以避免[疑難排解](Troubleshooting)中列出的問題。它們依主題分組。版面設定的建議適用於所有人，安全性和效能方面的建議對正式環境中的應用程式最為重要。

## <a id="page-setup"></a>版面設定

- **同時傳入全部四個邊界**。傳給 `Docxi::Document.new` 的雜湊會取代預設值，`top`、`right`、`bottom` 和 `left` 中漏掉的任何一項都會寫成空值。
- **設定第一頁的頁首和頁尾**。docxi 一律開啟獨立的第一頁。如果第 1 頁需要頁首或頁尾，還要呼叫 `add_first_header` 和 `add_first_footer`。

## <a id="text"></a>文字

- **開頭或結尾有空格時使用 `space: true`**。否則 Word 會捨棄這些空格，兩個文字片段的字就會黏在一起。
- **將共用的格式放在段落上**。傳給 `p` 的選項（例如 `size:` 或 `font:`）會套用到其中的每個文字片段，需要重複的內容就更少。
- **分行用 `p.br`，分頁用 `document.br(page: true)`**。沒有 `page: true` 的 `document.br` 會產生無效的 XML。

## <a id="tables"></a>表格

- **在列或儲存格上設定框線，而不是在表格上**。在 0.1.0 中，表格層級的 `borders:` 會套用到每個儲存格，而不是表格的外框。
- **為每一欄提供 `columns_width`**。儲存格會從中取得寬度，表格就能維持你預期的版面配置。

## <a id="images"></a>圖片

- **使用 PNG 檔案**。每張圖片都以 PNG 組件儲存。
- **用 `style:` 設定內嵌圖片的尺寸**，例如 `style: { width: "120pt", height: "60pt" }`。在 0.1.0 中，這是唯一能如預期運作的圖片尺寸設定方式。
- **如果文件可能渲染多次，請傳入路徑，而不是已開啟的 `File`**。IO 圖片會在 `render` 期間讀取，第二次渲染會得到空白圖片。
- **0.2.0 之前的頁首圖片、表格儲存格圖片和定位圖片**。它們的寬度和高度都以 `width * height * 14.92` EMU 計算，所以結果是正方形。若要得到邊長為 `S` 英吋的正方形圖片，請將 `width` 和 `height` 都設為 `Math.sqrt(S * 914400 / 14.92).round`，一英吋約為 248。

## <a id="security"></a>安全性

- **不要將不受信任的字串傳給 `add_media`**。傳入路徑會讓 docxi 讀取伺服器上的該檔案，傳入 URL 會讓你的伺服器下載它。請傳入你自己開啟的 IO，或先檢查該值。請參閱[安全性](Security)。
- **文字會自動跳脫**。傳給 `text` 的字串會以 XML 文字寫入，因此使用者輸入無法改變文件結構。

## <a id="performance-and-reliability"></a>效能與可靠性

- **在背景工作中建立大型文件**。呼叫 `render` 時，整份文件都會在記憶體中建立。請參閱[整合](Integrations)。
- **只渲染一次，然後重複使用位元組內容**。如果需要多次傳送或儲存文件，請保留 `word.render.string`。
- **用 `gem "docxi", "~> 0.1"` 鎖定相容的版本**。
- **測試你產生的內容**。`.docx` 是由 XML 檔案組成的 zip 壓縮檔，所以測試可以將它解壓縮並檢查 XML。docxi 自己的測試就是這樣做的。請參閱[架構](Architecture)。

## <a id="related-pages"></a>相關頁面

- [功能](Features)
- [設定](Configuration)
- [疑難排解](Troubleshooting)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
