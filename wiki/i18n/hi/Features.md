<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi की सुविधाएं: पैराग्राफ, टेक्स्ट, टेबल, इमेज, हेडर और फ़ुटर

यह पेज हर उस एलिमेंट की सूची देता है जिसे docxi किसी Word डॉक्यूमेंट में जोड़ सकता है, और हर एलिमेंट के विकल्प बताता है। उदाहरण इस सेटअप पर आधारित हैं:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

जिन विकल्पों की ज़रूरत न हो, उन्हें छोड़ा जा सकता है। रंग `#` के बिना हेक्स मान होते हैं, जैसे `"8B0000"`। यूनिट और पेज सेटअप के लिए [कॉन्फ़िगरेशन](Configuration) देखें।

## <a id="paragraphs"></a>पैराग्राफ

`document.p(options) { |p| ... }` एक पैराग्राफ जोड़ता है।

| विकल्प | मान | असर |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | अलाइनमेंट। `"both"` टेक्स्ट को जस्टिफ़ाई करता है |
| `fill` | हेक्स रंग | पैराग्राफ और उसके टेक्स्ट के पीछे शेडिंग |
| `left`, `right` | पॉइंट का बीसवां हिस्सा | इंडेंट। दोनों सेट करें, क्योंकि अकेला एक विकल्प अनदेखा हो जाता है |
| `ul` | `true` | पैराग्राफ को बुलेट सूची का आइटम बनाता है |
| `bottom` | `true` | पैराग्राफ के नीचे बॉर्डर बनाता है |

`size:`, `bold:` या `color:` जैसे टेक्स्ट विकल्प भी `p` को दिए जा सकते हैं। वे पैराग्राफ के हर टेक्स्ट रन के डिफ़ॉल्ट बन जाते हैं।

ब्लॉक के अंदर आप `text`, `br`, `tab`, `image`, `hyperlink` और `frame` कॉल कर सकते हैं। `document.text(string, options)` एक ही रन वाले पैराग्राफ का शॉर्टकट है।

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>टेक्स्ट

`p.text(string, options)` टेक्स्ट का एक रन (run) जोड़ता है।

| विकल्प | मान | असर |
|---|---|---|
| `bold` | `true` | बोल्ड |
| `italic` | `true` | इटैलिक |
| `underline` | `true` | सिंगल अंडरलाइन |
| `color` | हेक्स रंग | टेक्स्ट का रंग |
| `size` | पॉइंट | फ़ॉन्ट साइज़ |
| `font` | फ़ॉन्ट का नाम | फ़ॉन्ट, जैसे `"Georgia"` |
| `fill` | हेक्स रंग | टेक्स्ट के पीछे शेडिंग |
| `space` | `true` | टेक्स्ट की शुरुआत और आखिर के स्पेस बनाए रखता है |
| `br` | `true` | टेक्स्ट के बाद लाइन ब्रेक जोड़ता है |

`nil` स्ट्रिंग छोड़ दी जाती है।

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>लाइन ब्रेक, पेज ब्रेक और टैब

- `p.br`, या `p.br(times: 3)`, पैराग्राफ के अंदर लाइन ब्रेक जोड़ता है।
- `document.br(page: true)` नया पेज शुरू करता है।
- `p.tab`, या `p.tab(times: 2)`, पैराग्राफ के अंदर टैब कैरेक्टर जोड़ता है।
- `document.tab(page: true)` एक ऐसा पैराग्राफ जोड़ता है जिसमें सिर्फ़ एक टैब होता है।

`document.br` या `document.tab` को `page: true` के बिना कॉल न करें। ये टेक्स्ट रन को सीधे डॉक्यूमेंट बॉडी में रख देते हैं, जो मान्य WordprocessingML नहीं है।

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>टेबल

`document.table(options)` एक ब्लॉक लेता है, और उसके अंदर `table.tr(options)` पंक्तियां और `tr.tc(options)` सेल जोड़ता है।

**टेबल के विकल्प**

| विकल्प | मान | असर |
|---|---|---|
| `columns_width` | संख्याओं का array | कॉलम की चौड़ाई, जिसे 14.92 से गुणा करके पॉइंट के बीसवें हिस्से में बदला जाता है। यही सेल की डिफ़ॉल्ट चौड़ाई भी है |
| `height` | पॉइंट का बीसवां हिस्सा | हर पंक्ति की न्यूनतम ऊंचाई |
| `iwidth` | पॉइंट का बीसवां हिस्सा | बाएं मार्जिन से टेबल का इंडेंट |
| `style` | स्टाइल id | टेबल स्टाइल। docxi `TableNormal` और `TableCustom` परिभाषित करता है |
| `width` | संख्या | टेबल की पसंदीदा चौड़ाई, जो `auto` टाइप के साथ लिखी जाती है। `columns_width` न होने पर हर सेल को भी यही चौड़ाई मिलती है, 14.92 से गुणा करके |
| `align` | `"left"`, `"center"`, `"right"` | टेबल का अलाइनमेंट। `tc.text` से जोड़ा गया टेक्स्ट भी इसी तरह अलाइन होता है |
| `padding` | `true` | सेल के अंदर तय पैडिंग जोड़ता है |
| `borders` | Hash | ज्ञात समस्या: टेबल के बाहरी किनारों की जगह हर सेल पर लागू होता है |

**पंक्ति के विकल्प**: `height:` पंक्ति की न्यूनतम ऊंचाई तय करता है। `tr` को दिया गया सेल का कोई भी विकल्प उस पंक्ति के हर सेल पर लागू होता है। टेबल के विकल्प भी इसी तरह हर पंक्ति और हर सेल तक पहुंचते हैं।

**सेल के विकल्प**

| विकल्प | मान | असर |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | पॉइंट के आठवें हिस्से में बॉर्डर की मोटाई। `nil` का मतलब है उस तरफ़ कोई बॉर्डर नहीं |
| `fill` | हेक्स रंग | सेल की शेडिंग। बिना फ़िल वाले सेल सफ़ेद (`"FFFFFF"`) होते हैं |
| `valign` | `"top"`, `"center"`, `"bottom"` | वर्टिकल अलाइनमेंट |
| `colspan` | पूर्णांक | सेल कितने कॉलम में फैला है |
| `rowspan` | `true` | वर्टिकल मर्ज शुरू करता है |
| `merged` | `true` | ऊपर वाले सेल का वर्टिकल मर्ज जारी रखता है। इसका कंटेंट अनदेखा होता है |
| `width` | संख्या | सेल की चौड़ाई, 14.92 से गुणा की हुई |

**सेल का कंटेंट**

- `tc.text(string, options)`: सेल और पंक्ति के विकल्प, जैसे `bold:` या `align:`, टेक्स्ट पर भी लागू होते हैं।
- `tc.p(options) { |p| ... }`: सिर्फ़ अपने विकल्पों वाला पैराग्राफ।
- `tc.image(media, width:, height:)`: एक इमेज। साइज़ की समस्या के लिए [इमेज](#images) देखें।
- `tc.page_numbers`: पेज नंबर फ़ील्ड। ज्ञात समस्या: इसके आगे "GlobalOptions" टेक्स्ट लग जाता है।

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

## <a id="images"></a>इमेज

`media = document.add_media(source)` एक इमेज रजिस्टर करता है। `source` इनमें से कुछ भी हो सकता है:

- कोई IO ऑब्जेक्ट, जैसे `File` या `StringIO`
- फ़ाइल पाथ
- `http`, `https` या `ftp` URL

इमेज `render` कॉल करने पर पढ़ी जाती है और हमेशा PNG पार्ट के रूप में रखी जाती है, इसलिए PNG फ़ाइलें सुरक्षित विकल्प हैं।

**इनलाइन इमेज**: `p.image(media, options)` इमेज को टेक्स्ट की लाइन में रखता है।

| विकल्प | मान | असर |
|---|---|---|
| `style` | CSS प्रॉपर्टी का hash | साइज़, जैसे `{ width: "120pt", height: "60pt" }` |
| `title` | स्ट्रिंग | इमेज का शीर्षक |
| `fill_color` | VML रंग, जैसे `"#00FF00"` | इमेज के पीछे का फ़िल |
| `opacity` | स्ट्रिंग, जैसे `"50%"` | फ़िल की ओपैसिटी |
| `wrap` | `true` | इमेज के चारों ओर टेक्स्ट रैपिंग बंद करता है |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**पोज़िशन की गई इमेज**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` इमेज को कॉलम (`pH`) और पैराग्राफ (`pV`) से EMU में दिए गए ऑफ़सेट पर एंकर करता है। `pH`, `pV` और `wrap` तीनों देने ज़रूरी हैं, नहीं तो इमेज इनलाइन रखी जाती है।

**हेडर इमेज**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`। `wrap: "none"` रैपिंग बंद करता है। कोई भी दूसरा मान टेक्स्ट को इमेज के चारों ओर चौकोर रूप में रैप करता है।

**0.1.0 में ज्ञात समस्या**: पोज़िशन की गई इमेज, हेडर इमेज और टेबल सेल की इमेज अपनी चौड़ाई और ऊंचाई दोनों को `width * height * 14.92` EMU के रूप में निकालती हैं। ये हमेशा चौकोर बनती हैं और अक्सर बहुत छोटी होती हैं। `style:` से साइज़ की गई इनलाइन इमेज पर इसका असर नहीं पड़ता। वैकल्पिक उपाय के लिए [बेहतरीन तरीके](Best-Practices#images) देखें।

## <a id="headers-and-footers"></a>हेडर और फ़ुटर

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

- हेडर में `text`, `br`, `table`, `image` और `page_numbers` हो सकते हैं। फ़ुटर में `text`, `br`, `table` और `page_numbers` हो सकते हैं।
- `Header.new` या `Footer.new` को दिए गए विकल्प उनके टेक्स्ट पर लागू होते हैं।
- `add_header` और `add_footer` पेज 2 और उसके बाद के पेजों पर लागू होते हैं। पेज 1 पर हमेशा `add_first_header` और `add_first_footer` इस्तेमाल होते हैं। [कॉन्फ़िगरेशन](Configuration#first-page) देखें।

## <a id="page-numbers"></a>पेज नंबर

`page_numbers(name: "Page", align: "center")` पहले `name` का टेक्स्ट लिखता है और फिर एक PAGE फ़ील्ड। नंबर Arial 10 pt में, गहरे ग्रे रंग में होता है। `align` का डिफ़ॉल्ट `"right"` है।

ज्ञात समस्या: फ़ील्ड कोड में `\* MERGEFORMAT` का बैकस्लैश छूट जाता है।

## <a id="table-of-contents"></a>विषय-सूची

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`।
- पेज नंबर देने पर पंक्ति को `pos` पर दाईं ओर अलाइन टैब स्टॉप मिलता है (पॉइंट के बीसवें हिस्से में; 9360 का मतलब 6.5 इंच), जो `leader` कैरेक्टर से भरा होता है। लीडर के मान WordprocessingML से आते हैं, जैसे `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` और `"none"`।
- पेज नंबर न देने पर पंक्ति एक टैब के बाद का टेक्स्ट होती है।
- पेज नंबर वही होते हैं जो आप देते हैं। docxi Word का TOC फ़ील्ड नहीं बनाता और पेज की गिनती नहीं करता।

## <a id="hyperlinks"></a>हाइपरलिंक

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` एक रिलेशनशिप id लौटाता है, और `hyperlink(text, url, id)` लिंक रखता है। टारगेट रिलेशनशिप से आता है, इसलिए जिस URL की ज़रूरत है उसके लिए लौटाई गई id ही दें। सिर्फ़ बाहरी लिंक सपोर्ट किए जाते हैं।

## <a id="frames-text-boxes"></a>फ़्रेम (टेक्स्ट बॉक्स)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- फ़्रेम VML टेक्स्ट बॉक्स होते हैं, जिनमें तय हल्का ग्रे, आधा पारदर्शी फ़िल और काली आउटलाइन होती है। टेक्स्ट इनके ऊपर और नीचे रहता है।
- फ़्रेम के अंदर आप `text`, `br`, `tab`, `image` और `frame` कॉल कर सकते हैं।
- `spacing: true` टेक्स्ट से पहले जगह जोड़ता है।

## <a id="output"></a>आउटपुट

`word.render` पूरे `.docx` पैकेज वाला एक `StringIO` लौटाता है। आर्काइव Zip64 रिकॉर्ड के बिना लिखा जाता है, इसलिए Google Docs और LibreOffice के पुराने वर्शन इसे खोल सकते हैं। [आर्किटेक्चर](Architecture) पैकेज के हर पार्ट की सूची देता है।

## <a id="related-pages"></a>संबंधित पेज

- [उपयोग के उदाहरण](Usage-Examples)
- [API संदर्भ](API-Reference)
- [रोडमैप](Roadmap), जहां 0.2.0 के लिए तय ज्ञात समस्याएं हैं

---

[← विकी होम](Home) · [README](https://github.com/hammad417/docxi#readme)
