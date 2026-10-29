<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# डेवलपमेंट सेटअप: docxi के टेस्ट और टूल चलाना

यह पेज docxi की वर्किंग कॉपी चलाने में मदद करता है, और टेस्ट, सैंपल डॉक्यूमेंट और CI सेटअप के बारे में बताता है।

## <a id="get-the-code"></a>कोड लें

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

डेवलपमेंट के लिए 2.5 या उससे नया कोई भी Ruby चलता है। डेवलपमेंट डिपेंडेंसी minitest और rake हैं।

## <a id="rake-tasks"></a>Rake टास्क

| कमांड | क्या करता है |
|---|---|
| `bundle exec rake test` | टेस्ट सूट चलाता है (डिफ़ॉल्ट टास्क) |
| `bundle exec rake samples` | सैंपल डॉक्यूमेंट `tmp/samples/*.docx` में लिखता है, ताकि उन्हें Word, LibreOffice या Google Docs में खोला जा सके |
| `bundle exec rake golden` | मौजूदा कोड से `test/golden` फिर से लिखता है। इसे सिर्फ़ आउटपुट में जानबूझकर बदलाव के बाद इस्तेमाल करें |
| `bundle exec rake build` | gem को `pkg/` में बनाता है |

## <a id="running-the-tests-with-rails-loaded"></a>Rails लोड करके टेस्ट चलाना

`gemfiles/` फ़ोल्डर में हर Rails वर्शन के लिए एक Gemfile है, साथ में rubyzip और nokogiri के पुराने वर्शन के लिए कुछ और:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` टेस्ट से पहले पूरा फ़्रेमवर्क लोड करता है। यह docxi के `blank?` नियमों की तुलना ActiveSupport के नियमों से भी करता है।

## <a id="what-the-tests-cover"></a>टेस्ट क्या कवर करते हैं

| फ़ाइल | जांचती है |
|---|---|
| `test/test_golden_output.rb` | दो सैंपल डॉक्यूमेंट का हर पार्ट वैसा ही है जैसा 0.0.6 ने बनाया था |
| `test/test_package.rb` | XML सही तरह बना है, रिलेशनशिप मौजूद पार्ट की ओर इशारा करते हैं, हर पार्ट का कंटेंट टाइप है, और कोई Zip64 रिकॉर्ड नहीं है |
| `test/test_media.rb` | पाथ, IO ऑब्जेक्ट और URL से इमेज, और यह कि `\|` वाली स्ट्रिंग कभी कमांड की तरह नहीं चलती |
| `test/test_fixed_features.rb` | `tab`, `frame` और हेडर के पेज नंबर |
| `test/test_util.rb` | `Docxi::Util.blank?` ActiveSupport के नियमों का पालन करता है |
| `test/test_rails_independence.rb` | docxi बिना ActiveSupport के चलता है |
| `test/test_warnings.rb` | रेंडर करने पर `lib/` से कोई डिप्रिकेशन चेतावनी नहीं आती |

सैंपल डॉक्यूमेंट `test/support/sample_documents.rb` में हैं, और `test/support/package.rb` किसी `.docx` को सीधे उसके zip ढांचे से पढ़ता है।

## <a id="continuous-integration"></a>कंटीन्यूअस इंटीग्रेशन

`.github/workflows/ci.yml` `master` पर हर पुश और हर पुल रिक्वेस्ट पर चलता है। यह कवर करता है:

- बिना Rails के Ruby 2.5, 2.6 और 4.0
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) और 8.1 (4.0)
- Rails `main`, जिसके फ़ेल होने की अनुमति है
- rubyzip 1.3 और 2.4, और nokogiri 1.11
- एक जॉब जो LibreOffice से सैंपल डॉक्यूमेंट को PDF में बदलता है

रिलीज़ की जानकारी [डिप्लॉयमेंट](Deployment) में है।

## <a id="related-pages"></a>संबंधित पेज

- [योगदान](Contributing)
- [आर्किटेक्चर](Architecture)
- [डिप्लॉयमेंट](Deployment)

---

[← विकी होम](Home) · [README](https://github.com/hammad417/docxi#readme)
