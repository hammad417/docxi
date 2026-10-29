<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 功能：段落、文字、表格、圖片、頁首和頁尾

本頁列出 docxi 可以加入 Word 文件的每種元素，以及每種元素支援的選項。範例都以下列設定為基礎：

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

不需要的選項可以省略。色彩是不含 `#` 的十六進位值，例如 `"8B0000"`。單位和版面設定請參閱[設定](Configuration)。

## <a id="paragraphs"></a>段落

`document.p(options) { |p| ... }` 加入一個段落。

| 選項 | 值 | 效果 |
|---|---|---|
| `align` | `"left"`、`"center"`、`"right"`、`"both"` | 對齊方式。`"both"` 表示左右對齊 |
| `fill` | 十六進位色彩 | 段落及其文字後方的網底 |
| `left`、`right` | 二十分之一點 | 縮排。兩者要同時設定，只設定其中一個會被忽略 |
| `ul` | `true` | 將段落設為項目符號清單項目 |
| `bottom` | `true` | 在段落下方繪製框線 |

`size:`、`bold:`、`color:` 等文字選項也可以傳給 `p`，它們會成為段落中每個文字片段的預設值。

在區塊（block）中可以呼叫 `text`、`br`、`tab`、`image`、`hyperlink` 和 `frame`。`document.text(string, options)` 是只含一個文字片段之段落的簡寫。

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>文字

`p.text(string, options)` 加入一個文字片段（run）。

| 選項 | 值 | 效果 |
|---|---|---|
| `bold` | `true` | 粗體 |
| `italic` | `true` | 斜體 |
| `underline` | `true` | 單底線 |
| `color` | 十六進位色彩 | 文字色彩 |
| `size` | 點 | 字型大小 |
| `font` | 字型名稱 | 字型，例如 `"Georgia"` |
| `fill` | 十六進位色彩 | 文字後方的網底 |
| `space` | `true` | 保留文字開頭和結尾的空格 |
| `br` | `true` | 在文字後加入分行 |

`nil` 字串會被略過。

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>分行、分頁和定位字元

- `p.br` 或 `p.br(times: 3)` 在段落內加入分行。
- `document.br(page: true)` 開始新的一頁。
- `p.tab` 或 `p.tab(times: 2)` 在段落內加入定位字元。
- `document.tab(page: true)` 加入一個只含一個定位字元的段落。

不要在沒有 `page: true` 的情況下呼叫 `document.br` 或 `document.tab`。它們會把文字片段直接放進文件本文，這不是有效的 WordprocessingML。

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>表格

`document.table(options)` 接受一個區塊，在區塊中以 `table.tr(options)` 加入列，以 `tr.tc(options)` 加入儲存格。

**表格選項**

| 選項 | 值 | 效果 |
|---|---|---|
| `columns_width` | 數字陣列 | 欄寬，乘以 14.92 換算成二十分之一點。也是儲存格的預設寬度 |
| `height` | 二十分之一點 | 每一列的最小高度 |
| `iwidth` | 二十分之一點 | 表格相對於左邊界的縮排 |
| `style` | 樣式 ID | 表格樣式。docxi 定義了 `TableNormal` 和 `TableCustom` |
| `width` | 數字 | 表格的慣用寬度，類型寫為 `auto`。未設定 `columns_width` 時，每個儲存格也會使用這個寬度，並乘以 14.92 |
| `align` | `"left"`、`"center"`、`"right"` | 表格對齊方式。以 `tc.text` 加入的文字也會以相同方式對齊 |
| `padding` | `true` | 在儲存格內加入固定的內距 |
| `borders` | 雜湊 | 已知問題：會套用到每個儲存格，而不是表格的外框 |

**列選項**：`height:` 設定該列的最小高度。傳給 `tr` 的任何儲存格選項都會套用到該列的每個儲存格。表格選項也會以相同方式傳給每一列和每個儲存格。

**儲存格選項**

| 選項 | 值 | 效果 |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | 框線粗細，單位為八分之一點。`nil` 表示該邊沒有框線 |
| `fill` | 十六進位色彩 | 儲存格網底。沒有設定填滿的儲存格為白色（`"FFFFFF"`） |
| `valign` | `"top"`、`"center"`、`"bottom"` | 垂直對齊 |
| `colspan` | 整數 | 儲存格跨越的欄數 |
| `rowspan` | `true` | 開始垂直合併 |
| `merged` | `true` | 延續上方儲存格的垂直合併，其內容會被忽略 |
| `width` | 數字 | 儲存格寬度，乘以 14.92 |

**儲存格內容**

- `tc.text(string, options)`：儲存格和列的選項（例如 `bold:` 或 `align:`）也會套用到文字。
- `tc.p(options) { |p| ... }`：只使用自身選項的段落。
- `tc.image(media, width:, height:)`：一張圖片。尺寸問題請參閱下方的[圖片](#images)。
- `tc.page_numbers`：頁碼功能變數。已知問題：前面會多出文字「GlobalOptions」。

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>圖片

`media = document.add_media(source)` 註冊一張圖片。`source` 可以是：

- IO 物件，例如 `File` 或 `StringIO`
- 檔案路徑
- `http`、`https` 或 `ftp` URL

圖片會在呼叫 `render` 時讀取，並一律以 PNG 組件儲存，所以 PNG 檔案是最保險的選擇。

**內嵌圖片**：`p.image(media, options)` 將圖片放在文字行中。

| 選項 | 值 | 效果 |
|---|---|---|
| `style` | CSS 屬性的雜湊 | 尺寸，例如 `{ width: "120pt", height: "60pt" }` |
| `title` | 字串 | 圖片標題 |
| `fill_color` | VML 色彩，例如 `"#00FF00"` | 圖片後方的填滿 |
| `opacity` | 字串，例如 `"50%"` | 填滿的不透明度 |
| `wrap` | `true` | 關閉圖片周圍的文繞圖 |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**定位圖片**：`p.image(media, pH:, pV:, wrap: true, width:, height:)` 以 EMU 為單位，將圖片錨定在相對於欄（`pH`）和段落（`pV`）的位移處。`pH`、`pV` 和 `wrap` 必須同時提供，否則圖片會以內嵌方式放置。

**頁首圖片**：`h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`。`wrap: "none"` 會關閉文繞圖，其他任何值都會讓文字以矩形方式環繞圖片。

**0.1.0 的已知問題**：定位圖片、頁首圖片和表格儲存格圖片的寬度和高度都以 `width * height * 14.92` EMU 計算。它們一律是正方形，而且通常很小。以 `style:` 設定尺寸的內嵌圖片不受影響。因應措施請參閱[最佳做法](Best-Practices#images)。

## <a id="headers-and-footers"></a>頁首和頁尾

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- 頁首可以包含 `text`、`br`、`table`、`image` 和 `page_numbers`。頁尾可以包含 `text`、`br`、`table` 和 `page_numbers`。
- 傳給 `Header.new` 或 `Footer.new` 的選項會套用到其中的文字。
- `add_header` 和 `add_footer` 用於第 2 頁之後的頁面。第 1 頁一律使用 `add_first_header` 和 `add_first_footer`。請參閱[設定](Configuration#first-page)。

## <a id="page-numbers"></a>頁碼

`page_numbers(name: "Page", align: "center")` 先寫入 `name` 文字，再寫入一個 PAGE 功能變數。頁碼使用 Arial 10 點、深灰色。`align` 預設為 `"right"`。

已知問題：功能變數代碼中 `\* MERGEFORMAT` 缺少反斜線。

## <a id="table-of-contents"></a>目錄

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`。
- 提供頁碼時，該項目會在 `pos` 處取得一個靠右對齊的定位點（單位為二十分之一點，9360 即 6.5 英吋），並以 `leader` 字元填滿。前置字元的值來自 WordprocessingML，例如 `"dot"`、`"hyphen"`、`"underscore"`、`"middleDot"` 和 `"none"`。
- 未提供頁碼時，該項目是定位字元後面的文字。
- 頁碼就是你傳入的值。docxi 不會建立 Word 的 TOC 功能變數，也不會計算頁碼。

## <a id="hyperlinks"></a>超連結

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` 傳回一個關聯 ID，`hyperlink(text, url, id)` 負責放置連結。連結目標來自關聯，所以要傳入為目標 URL 傳回的那個 ID。只支援外部連結。

## <a id="frames-text-boxes"></a>框架（文字方塊）

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- 框架是 VML 文字方塊，帶有固定的淺灰色半透明填滿和黑色外框。文字會排在它的上方和下方。
- 在框架中可以呼叫 `text`、`br`、`tab`、`image` 和 `frame`。
- `spacing: true` 會在文字前加入間距。

## <a id="output"></a>輸出

`word.render` 傳回一個包含完整 `.docx` 封裝的 `StringIO`。壓縮檔中不含 Zip64 記錄，因此 Google Docs 和較舊的 LibreOffice 版本可以開啟它。[架構](Architecture)列出了封裝中的每個組件。

## <a id="related-pages"></a>相關頁面

- [使用範例](Usage-Examples)
- [API 參考](API-Reference)
- [路線圖](Roadmap)：預計在 0.2.0 修正的已知問題

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
