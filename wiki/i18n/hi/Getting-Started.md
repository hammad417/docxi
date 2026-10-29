<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# docxi के साथ शुरुआत: Ruby में आपका पहला Word डॉक्यूमेंट

यह गाइड आपको एक खाली फ़ोल्डर से फ़ॉर्मैट की हुई `.docx` फ़ाइल तक ले जाती है। इसमें माना गया है कि आप Ruby की बुनियादी बातें जानते हैं। आपको Rails या Microsoft Word की ज़रूरत नहीं है।

## <a id="1-install-the-gem"></a>1. gem इंस्टॉल करें

```sh
gem install docxi
```

अगर प्रोजेक्ट में Gemfile है, तो इसके बजाय `gem "docxi", "~> 0.1"` जोड़ें और `bundle install` चलाएं। पूरी जानकारी [इंस्टॉलेशन](Installation) में है।

## <a id="2-create-a-document"></a>2. डॉक्यूमेंट बनाएं

इसे `hello.rb` नाम से सेव करें:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

`ruby hello.rb` चलाएं (Bundler प्रोजेक्ट में `bundle exec ruby hello.rb`) और `hello.docx` खोलें।

## <a id="3-what-each-line-does"></a>3. हर लाइन क्या करती है

- `Docxi::Document.new` डिफ़ॉल्ट पेज मार्जिन वाला एक खाली डॉक्यूमेंट बनाता है।
- `word.document` डॉक्यूमेंट की बॉडी है। आप जो कुछ भी जोड़ते हैं, वह इसी के ज़रिए जाता है।
- `document.p` एक पैराग्राफ जोड़ता है। `p` को दिए गए विकल्प, जैसे `align:`, पैराग्राफ और उसके अंदर के टेक्स्ट दोनों पर लागू होते हैं।
- `p.text` टेक्स्ट का एक रन (run) जोड़ता है। `size:` पॉइंट में होता है।
- `word.render` `.docx` पैकेज बनाता है और एक `StringIO` लौटाता है। उसके `.string` में फ़ाइल के बाइट होते हैं।

## <a id="4-add-a-list-and-a-table"></a>4. एक सूची और एक टेबल जोड़ें

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

दो बातों पर ध्यान दें:

- `space: true` किसी रन की शुरुआत या आखिर के स्पेस को बनाए रखता है। इसके बिना Word उस स्पेस को हटा देता है और शब्द आपस में चिपक जाते हैं।
- `tr` पर दिए गए विकल्प उस पंक्ति के हर सेल पर लागू होते हैं, इसलिए पूरी हेडर पंक्ति शेडेड और बोल्ड हो जाती है।

## <a id="5-where-to-go-next"></a>5. आगे क्या पढ़ें

- [उपयोग के उदाहरण](Usage-Examples): रिपोर्ट, इनवॉइस और पत्र के पूरे उदाहरण
- [सुविधाएं](Features): पैराग्राफ, टेक्स्ट, टेबल, इमेज, हेडर और फ़ुटर के सभी विकल्प
- [इंटीग्रेशन](Integrations): Rails controller से डॉक्यूमेंट भेजना
- [बेहतरीन तरीके](Best-Practices): आम गड़बड़ियों से पहले ही बचें

---

[← विकी होम](Home) · [README](https://github.com/hammad417/docxi#readme)
