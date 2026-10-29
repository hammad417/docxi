<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API संदर्भ: क्लास और मेथड

docxi एक Ruby लाइब्रेरी है, इसलिए यह पेज उसकी Ruby API की सूची देता है। इसकी कोई HTTP API नहीं है। हर मेथड के विकल्प [सुविधाएं](Features) में दिए गए हैं, और क्लास [आर्किटेक्चर](Architecture) में समझाई गई हैं।

## <a id="docxidocument"></a>Docxi::Document

| मेथड | लौटाता है | टिप्पणी |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` पेज मार्जिन हैं। [कॉन्फ़िगरेशन](Configuration) देखें |
| `#document` | `Docxi::Word::Document` | वह बॉडी जिसमें आप कंटेंट जोड़ते हैं |
| `#render` | `StringIO` | पूरा `.docx` पैकेज। `.string` बाइट देता है |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| मेथड | लौटाता है | जोड़ता है |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | एक पैराग्राफ |
| `text(string, options = {})` | Paragraph | टेक्स्ट के एक रन वाला पैराग्राफ |
| `br(page: true)` | Break | पेज ब्रेक |
| `tab(page: true)` | Tab | एक टैब वाला पैराग्राफ |
| `table(options = {}) { \|table\| }` | Table | एक टेबल |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | विषय-सूची की पंक्तियां |
| `frame(text, options = {})` | Paragraph | टेक्स्ट बॉक्स वाला पैराग्राफ |
| `add_media(io_path_or_url, options = {})` | Media | `image` से रखी जाने वाली इमेज |
| `add_hyperlink(url, options = {})` | String | `hyperlink` के लिए रिलेशनशिप id |
| `add_header(header)`, `add_first_header(header)` | Header | पेज 2 से आगे का हेडर, और पेज 1 का हेडर |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | पेज 2 से आगे का फ़ुटर, और पेज 1 का फ़ुटर |

## <a id="paragraph-docxiwordcontentsparagraph"></a>पैराग्राफ (`Docxi::Word::Contents::Paragraph`)

| मेथड | जोड़ता है |
|---|---|
| `text(string, options = {})` | टेक्स्ट का एक रन। पैराग्राफ के विकल्प इसके डिफ़ॉल्ट होते हैं |
| `br(times: 1)` | लाइन ब्रेक |
| `tab(times: 1)` | टैब कैरेक्टर |
| `image(media, options = {})` | इनलाइन या पोज़िशन की गई इमेज |
| `hyperlink(text, url, id)` | ऐसा लिंक जो `add_hyperlink` से मिली रिलेशनशिप `id` इस्तेमाल करता है |
| `frame(options = {}) { \|frame\| }` | एक टेक्स्ट बॉक्स |

## <a id="tables"></a>टेबल

| मेथड | जोड़ता है |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | एक पंक्ति। टेबल के विकल्प इसके डिफ़ॉल्ट होते हैं |
| `TableRow#tc(options = {}) { \|tc\| }` | एक सेल। पंक्ति के विकल्प इसके डिफ़ॉल्ट होते हैं |
| `TableCell#text(string, options = {})` | एक पैराग्राफ। सेल के विकल्प इसके डिफ़ॉल्ट होते हैं |
| `TableCell#p(options = {}) { \|p\| }` | अपने विकल्पों वाला पैराग्राफ |
| `TableCell#image(media, width:, height:)` | एक इमेज |
| `TableCell#page_numbers` | पेज नंबर फ़ील्ड |

## <a id="table-of-contents"></a>विषय-सूची

| मेथड | जोड़ता है |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | विषय-सूची की एक पंक्ति |

## <a id="headers-and-footers"></a>हेडर और फ़ुटर

| क्लास | मेथड |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>फ़्रेम

`Docxi::Word::Contents::Frame` अपने ब्लॉक के अंदर `text`, `br`, `tab`, `image` और `frame` स्वीकार करता है।

## <a id="constants-and-helpers"></a>कॉन्स्टेंट और हेल्पर

| नाम | मतलब |
|---|---|
| `Docxi::VERSION` | इंस्टॉल किया गया वर्शन, जैसे `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | आर्काइव में एक पार्ट लिखता है और उसका साइज़ पहले से सेट करता है। हर पार्ट अंदरूनी तौर पर इसे इस्तेमाल करता है |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | ActiveSupport के `blank?` नियमों वाले अंदरूनी हेल्पर |

## <a id="related-pages"></a>संबंधित पेज

- [सुविधाएं](Features)
- [उपयोग के उदाहरण](Usage-Examples)
- [आर्किटेक्चर](Architecture)

---

[← विकी होम](Home) · [README](https://github.com/hammad417/docxi#readme)
