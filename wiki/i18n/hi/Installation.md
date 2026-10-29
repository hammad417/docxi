<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# docxi इंस्टॉल करना: ज़रूरतें, Bundler सेटअप और अपग्रेड

docxi पूरी तरह Ruby में लिखा gem है। इसे इंस्टॉल करने पर सिर्फ़ दो डिपेंडेंसी आती हैं, nokogiri और rubyzip, और कुछ नहीं। किसी ऑफ़िस सॉफ़्टवेयर की ज़रूरत नहीं होती।

## <a id="requirements"></a>ज़रूरतें

| ज़रूरत | वर्शन |
|---|---|
| Ruby | 2.5 या उससे नया |
| nokogiri | `>= 1.8, < 2`, अपने-आप इंस्टॉल होता है |
| rubyzip | `>= 1.3, < 4`, अपने-आप इंस्टॉल होता है |
| Rails | ज़रूरी नहीं |

## <a id="install-with-bundler"></a>Bundler के साथ इंस्टॉल करें

gem को अपनी Gemfile में जोड़ें:

```ruby
gem "docxi", "~> 0.1"
```

फिर इसे इंस्टॉल करें:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Bundler के बिना इंस्टॉल करें

```sh
gem install docxi
```

## <a id="check-the-installation"></a>इंस्टॉलेशन जांचें

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Bundler प्रोजेक्ट में इसे `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` के रूप में चलाएं। यह इंस्टॉल किया गया वर्शन दिखाता है, जैसे `0.1.0`।

## <a id="upgrading-from-00x"></a>0.0.x से अपग्रेड करना

अगर आपके Gemfile में docxi किसी 0.0.x वर्ज़न पर तय है, जैसे `"~> 0.0.6"`, तो उस लाइन को `gem "docxi", "~> 0.1"` में बदलें। फिर `bundle update docxi` चलाएं। कोड में कोई बदलाव ज़रूरी नहीं है, और 0.1.0 वही डॉक्यूमेंट बनाता है जो 0.0.6 बनाता था: टेस्ट सूट आउटपुट की तुलना 0.0.6 से बनी फ़ाइलों से करता है।

अंदर क्या बदला:

- nokogiri और rubyzip अब ऊपर दी गई रेंज का कोई भी वर्शन हो सकते हैं, इसलिए docxi मौजूदा Rails वर्शन के साथ इंस्टॉल हो जाता है।
- docxi को अब ActiveSupport लोड होने की ज़रूरत नहीं है।
- आर्काइव Zip64 रिकॉर्ड के बिना लिखे जाते हैं।
- इमेज `Kernel#open` के बिना लोड होती हैं, इसलिए Ruby 3 पर URL काम करते हैं।
- `tab`, `frame` और हेडर के पेज नंबर काम करते हैं।

2.5 से पुराने Ruby पर चलने वाले ऐप्लिकेशन को 0.0.6 ही मिलता रहेगा, क्योंकि 0.1.0 में Ruby 2.5 न्यूनतम वर्शन के रूप में दर्ज है। [बदलाव का लॉग](Changelog) हर बदलाव की सूची देता है।

## <a id="versions-tested-in-ci"></a>CI में टेस्ट किए गए वर्शन

हर Rails वर्शन पूरे फ़्रेमवर्क को docxi के साथ लोड करके चलाया जाता है।

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI ये भी चलाता है:

- Ruby 2.5, 2.6 और 4.0 पर बिना Rails के टेस्ट
- rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) और nokogiri 1.11 (Ruby 2.5) के साथ टेस्ट
- यह जांच कि LibreOffice सैंपल डॉक्यूमेंट को बदल सकता है

## <a id="platform-notes"></a>प्लेटफ़ॉर्म से जुड़ी बातें

- nokogiri आम Linux, macOS और Windows प्लेटफ़ॉर्म के लिए प्रीकंपाइल्ड gem देता है। जहां कोई मेल नहीं खाता, वहां यह सोर्स से कंपाइल होता है और इसके लिए C कंपाइलर चाहिए। [nokogiri इंस्टॉलेशन गाइड](https://nokogiri.org/tutorials/installing_nokogiri.html) देखें।
- CI, Ubuntu 24.04 पर चलता है, और टेस्ट macOS पर भी चलाए गए हैं। Windows पर टेस्ट नहीं किया जाता।

## <a id="if-installation-fails"></a>अगर इंस्टॉलेशन फ़ेल हो

nokogiri या rubyzip पर Bundler के टकराव का मतलब लगभग हमेशा यह होता है कि पुराना docxi 0.0.6 रिज़ॉल्व हो रहा है। [समस्या निवारण](Troubleshooting#installation) देखें।

## <a id="related-pages"></a>संबंधित पेज

- [शुरुआत करें](Getting-Started)
- [कॉन्फ़िगरेशन](Configuration)
- [डिप्लॉयमेंट](Deployment)

---

[← विकी होम](Home) · [README](https://github.com/hammad417/docxi#readme)
