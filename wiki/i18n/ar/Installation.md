<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# تثبيت docxi: المتطلبات وإعداد Bundler والترقية

مكتبة docxi حزمة مكتوبة بلغة Ruby بالكامل. يجلب تثبيتها اعتماديتين فقط، هما nokogiri وrubyzip، ولا شيء غيرهما. لا حاجة إلى أي برنامج مكتبي.

## <a id="requirements"></a>المتطلبات

| المتطلب | الإصدار |
|---|---|
| لغة Ruby | الإصدار 2.5 أو أحدث |
| مكتبة nokogiri | <span dir="ltr">`>= 1.8, < 2`</span>، تُثبَّت تلقائيًا |
| مكتبة rubyzip | <span dir="ltr">`>= 1.3, < 4`</span>، تُثبَّت تلقائيًا |
| إطار Rails | غير مطلوب |

## <a id="install-with-bundler"></a>التثبيت باستخدام Bundler

أضف الحزمة إلى ملف Gemfile:

```ruby
gem "docxi", "~> 0.1"
```

ثم ثبّتها:

```sh
bundle install
```

## <a id="install-without-bundler"></a>التثبيت دون Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>التحقق من التثبيت

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

في مشروع يستخدم Bundler، شغّله بالشكل <span dir="ltr">`bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`</span>. يطبع الأمر الإصدار المثبّت، مثل <span dir="ltr">`0.1.0`</span>.

## <a id="upgrading-from-00x"></a>الترقية من 0.0.x

إذا كان ملف Gemfile يثبّت docxi على إصدار 0.0.x، مثل <span dir="ltr">`"~> 0.0.6"`</span>، فغيّر السطر إلى <span dir="ltr">`gem "docxi", "~> 0.1"`</span>. ثم شغّل <span dir="ltr">`bundle update docxi`</span>. لا حاجة إلى أي تغيير في الشيفرة، فالإصدار 0.1.0 يولّد المستندات نفسها التي يولّدها 0.0.6، إذ تقارن مجموعة الاختبارات المخرجات بملفات ولّدها الإصدار 0.0.6.

ما الذي تغيّر داخليًا:

- يمكن الآن أن تكون nokogiri وrubyzip بأي إصدار ضمن النطاقات السابقة، لذا تُثبَّت docxi إلى جانب إصدارات Rails الحالية.
- لم تعد docxi تحتاج إلى تحميل ActiveSupport.
- تُكتب الأرشيفات دون سجلات Zip64.
- تُحمَّل الصور دون <span dir="ltr">`Kernel#open`</span>، لذا تعمل عناوين URL على Ruby 3.
- تعمل الآن <span dir="ltr">`tab`</span> و<span dir="ltr">`frame`</span> وأرقام الصفحات في رؤوس الصفحات.

التطبيقات التي تعمل بإصدار Ruby أقدم من 2.5 تستمر في الحصول على 0.0.6، لأن الإصدار 0.1.0 يحدد Ruby 2.5 حدًا أدنى. يسرد [سجل التغييرات](Changelog) كل التغييرات.

## <a id="versions-tested-in-ci"></a>الإصدارات المختبرة في CI

يعمل كل إصدار من Rails مع تحميل إطار العمل كاملًا إلى جانب docxi.

| إصدار Ruby | إصدار Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

تشغّل CI أيضًا:

- الاختبارات دون Rails على Ruby 2.5 و2.6 و4.0
- اختبارات مع rubyzip 1.3 ‏(Ruby 2.5) وrubyzip 2.4 ‏(Ruby 3.3) وnokogiri 1.11 ‏(Ruby 2.5)
- فحصًا للتأكد من أن LibreOffice يستطيع تحويل المستندات النموذجية

## <a id="platform-notes"></a>ملاحظات حول المنصات

- توفّر nokogiri حزمًا مُصرَّفة مسبقًا لمنصات Linux وmacOS وWindows الشائعة. وحين لا تتوفر حزمة مطابقة، تُصرَّف من الشيفرة المصدرية وتحتاج إلى مصرّف C. راجع [دليل تثبيت nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- تعمل CI على Ubuntu 24.04، كما شُغّلت الاختبارات على macOS. أما Windows فلم يُختبر.

## <a id="if-installation-fails"></a>إذا فشل التثبيت

تعارضات Bundler حول nokogiri أو rubyzip تعني في الغالب أن الإصدار القديم docxi 0.0.6 هو الذي يجري حلّه. راجع [استكشاف الأخطاء وإصلاحها](Troubleshooting#installation).

## <a id="related-pages"></a>صفحات ذات صلة

- [البدء](Getting-Started)
- [الإعدادات](Configuration)
- [النشر](Deployment)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
