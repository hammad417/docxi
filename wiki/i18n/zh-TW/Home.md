<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi wiki：用 Ruby 產生 Word（.docx）文件

docxi 是一個 Ruby gem，用程式碼產生 Microsoft Word `.docx` 檔案。它支援段落、帶格式的文字、表格、圖片、頁首和頁尾、頁碼、超連結以及簡單的目錄。它不需要 Microsoft Word、LibreOffice 或範本，在一般的 Ruby 程式和 Ruby on Rails 應用程式中都能使用。

[README](https://github.com/hammad417/docxi#readme) 是簡短的介紹。本 wiki 涵蓋所有選項，提供可直接複製的範例，說明 docxi 的內部運作方式，並在出問題時提供協助。

目前版本：**0.1.1**，適用於 Ruby 2.5 以上版本。請參閱[變更紀錄](Changelog)。

## <a id="start-here"></a>從這裡開始

**第一次使用 docxi**

- [快速入門](Getting-Started)：安裝 gem 並建立第一份文件
- [安裝](Installation)：系統需求、Bundler，以及從 0.0.x 升級
- [使用範例](Usage-Examples)：報告、發票、信件和由資料產生的表格，可直接複製

**建立文件**

- [功能](Features)：每種元素及其支援的選項
- [設定](Configuration)：邊界、版面設定和單位
- [API 參考](API-Reference)：類別和方法一覽
- [最佳做法](Best-Practices)：避開常見意外的習慣

**出問題時**

- [疑難排解](Troubleshooting)：錯誤訊息，以及顯示不正常的文件
- [常見問題](FAQ)：常見問題的簡短解答
- [安全性](Security)：回報問題以及處理不受信任的輸入

**深入了解**

- [架構](Architecture)：Ruby 呼叫如何變成 `.docx` 封裝
- [整合](Integrations)：Rails 控制器、Rack 應用程式和背景工作
- [部署](Deployment)：在正式環境中執行 docxi，以及版本如何發布
- [參與貢獻](Contributing)和[開發環境設定](Development-Setup)
- [路線圖](Roadmap)、[變更紀錄](Changelog)和[致謝](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi 能做什麼，不能做什麼

docxi 用來建立新文件。它無法開啟或編輯現有的 `.docx`，無法填入 Word 範本，也無法執行合併列印。它沒有編號清單、標題樣式或自動目錄，頁面一律是直向的 US Letter。預計在 0.2.0 修正的已知問題列在[路線圖](Roadmap)上。

## <a id="credits"></a>致謝

docxi 由 [Phaedra Solutions](https://www.phaedrasolutions.com/) 的 [Hammad Maqbool](https://github.com/hammad417) 維護。它於 2016 年以 Timur Gabdrakipov 的 [docxer](https://github.com/gtimti/docxer) 為基礎起步。詳情請參閱[致謝](Credits)。

## <a id="languages"></a>語言

英文版是本 wiki 的參考版本。各語言譯本的連結位於側邊欄和每個頁面的頂端。

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [回報問題](https://github.com/hammad417/docxi/issues)
