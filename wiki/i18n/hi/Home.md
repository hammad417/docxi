<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi विकी: Ruby से Word (.docx) डॉक्यूमेंट बनाएं

docxi एक Ruby gem है जो कोड से Microsoft Word की `.docx` फ़ाइलें बनाता है। इसमें पैराग्राफ, फ़ॉर्मैट किया हुआ टेक्स्ट, टेबल, इमेज, हेडर और फ़ुटर, पेज नंबर, हाइपरलिंक और एक साधारण विषय-सूची मिलती है। इसके लिए Microsoft Word, LibreOffice या किसी टेम्पलेट की ज़रूरत नहीं होती, और यह साधारण Ruby के साथ-साथ Ruby on Rails ऐप्लिकेशन में भी काम करता है।

[README](https://github.com/hammad417/docxi#readme) छोटा परिचय है। यह विकी हर विकल्प को कवर करता है, कॉपी करने लायक तैयार उदाहरण देता है, बताता है कि docxi अंदर से कैसे काम करता है, और कुछ गड़बड़ होने पर मदद करता है।

मौजूदा रिलीज़: **0.1.1**, Ruby 2.5 या उससे नए वर्शन के लिए। [बदलाव का लॉग](Changelog) देखें।

## <a id="start-here"></a>यहां से शुरू करें

**docxi में नए हैं**

- [शुरुआत करें](Getting-Started): gem इंस्टॉल करें और अपना पहला डॉक्यूमेंट बनाएं
- [इंस्टॉलेशन](Installation): ज़रूरतें, Bundler, और 0.0.x से अपग्रेड
- [उपयोग के उदाहरण](Usage-Examples): डेटा से बनी रिपोर्ट, इनवॉइस, पत्र और टेबल, कॉपी करने के लिए तैयार

**डॉक्यूमेंट बनाना**

- [सुविधाएं](Features): हर एलिमेंट और उसके विकल्प
- [कॉन्फ़िगरेशन](Configuration): मार्जिन, पेज सेटअप और यूनिट
- [API संदर्भ](API-Reference): क्लास और मेथड एक नज़र में
- [बेहतरीन तरीके](Best-Practices): ऐसी आदतें जो आम गड़बड़ियों से बचाती हैं

**जब कुछ गड़बड़ हो**

- [समस्या निवारण](Troubleshooting): एरर मैसेज और ऐसे डॉक्यूमेंट जो सही नहीं दिखते
- [अक्सर पूछे जाने वाले सवाल](FAQ): आम सवालों के छोटे जवाब
- [सुरक्षा](Security): समस्याओं की रिपोर्ट करना और अविश्वसनीय इनपुट को संभालना

**और गहराई में**

- [आर्किटेक्चर](Architecture): Ruby कॉल `.docx` पैकेज कैसे बनती हैं
- [इंटीग्रेशन](Integrations): Rails controller, Rack ऐप और बैकग्राउंड जॉब
- [डिप्लॉयमेंट](Deployment): प्रोडक्शन में docxi चलाना, और रिलीज़ कैसे प्रकाशित होती हैं
- [योगदान](Contributing) और [डेवलपमेंट सेटअप](Development-Setup)
- [रोडमैप](Roadmap), [बदलाव का लॉग](Changelog) और [श्रेय](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi क्या करता है, और क्या नहीं

docxi नए डॉक्यूमेंट बनाता है। यह किसी मौजूदा `.docx` को खोल या एडिट नहीं कर सकता, Word टेम्पलेट नहीं भर सकता और मेल मर्ज नहीं चला सकता। इसमें नंबर वाली सूचियां, हेडिंग स्टाइल या अपने-आप बनने वाली विषय-सूची नहीं है, और पेज हमेशा पोर्ट्रेट में US Letter होते हैं। 0.2.0 के लिए तय ज्ञात समस्याएं [रोडमैप](Roadmap) पर दी गई हैं।

## <a id="credits"></a>श्रेय

docxi का रखरखाव [Phaedra Solutions](https://www.phaedrasolutions.com/) में [Hammad Maqbool](https://github.com/hammad417) द्वारा किया जाता है। इसकी शुरुआत 2016 में Timur Gabdrakipov के [docxer](https://github.com/gtimti/docxer) से हुई थी। विवरण के लिए [श्रेय](Credits) देखें।

## <a id="languages"></a>भाषाएं

इस विकी का अंग्रेज़ी वर्शन संदर्भ वर्शन है। अनुवादों के लिंक साइडबार में और हर पेज के ऊपर दिए गए हैं।

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [समस्या की रिपोर्ट करें](https://github.com/hammad417/docxi/issues)
