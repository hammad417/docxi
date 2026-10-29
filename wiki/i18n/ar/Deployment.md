<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# تشغيل docxi في بيئة الإنتاج ونشر الإصدارات

مكتبة docxi مكتبة وليست خدمة، لذا لا يوجد ما يُنشر منها بمفردها. تتناول هذه الصفحة ما يهم عند نشر تطبيق يستخدم docxi، وكيف تُنشر الإصدارات الجديدة من docxi.

## <a id="shipping-an-app-that-uses-docxi"></a>نشر تطبيق يستخدم docxi

- **الاعتماديات.** يثبّت <span dir="ltr">`bundle install`</span> المكتبتين nokogiri وrubyzip. توفّر nokogiri حزمًا مُصرَّفة مسبقًا لمنصات Linux وmacOS وWindows الشائعة. أما على المنصات الأخرى، مثل بعض صور الحاويات المصغّرة، فتُصرَّف من الشيفرة المصدرية وتحتاج إلى مصرّف C. راجع [دليل تثبيت nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **لا حاجة إلى برامج مكتبية.** لا تحتاج الخوادم والحاويات إلى Microsoft Word أو LibreOffice.
- **الذاكرة.** تبني <span dir="ltr">`render`</span> المستند كاملًا في الذاكرة وتعيده ككائن <span dir="ltr">`StringIO`</span>. يزداد استهلاك الذاكرة بزيادة عدد الصفحات وحجم الصور.
- **الشبكة والملفات.** لا تقرأ docxi نظام الملفات إلا لمسارات الصور التي تمررها إلى <span dir="ltr">`add_media`</span>، ولا تستخدم الشبكة إلا لعناوين URL للصور التي تمررها. راجع [الأمان](Security).
- **المستندات الطويلة.** ابنِها في مهمة في الخلفية بدلًا من بنائها أثناء طلب ويب. راجع [التكامل](Integrations).

## <a id="how-docxi-releases-are-published"></a>كيف تُنشر إصدارات docxi

هذه الخطوات موجّهة إلى القائمين على الصيانة. تُنشر الإصدارات على RubyGems عبر GitHub Actions باستخدام النشر الموثوق (trusted publishing)، لذا لا يُخزَّن أي مفتاح API في المستودع.

1. حدّث <span dir="ltr">`lib/docxi/version.rb`</span> و<span dir="ltr">`CHANGELOG.md`</span>، وادمج التغييرات في <span dir="ltr">`master`</span>.
2. ضع على الإيداع (commit) وسمًا باسم الإصدار مسبوقًا بالحرف <span dir="ltr">`v`</span>، مثل <span dir="ltr">`v0.1.1`</span>، ثم ادفع الوسم.
3. يبني <span dir="ltr">`.github/workflows/release.yml`</span> الحزمة ويدفعها إلى RubyGems.

تتبع الإصدارات التجريبية الخطوات نفسها برقم إصدار مثل <span dir="ltr">`0.2.0.rc1`</span>. لا يثبّت Bundler الإصدارات التجريبية إلا عند طلبها صراحةً. تعمل CI ‏(<span dir="ltr">`.github/workflows/ci.yml`</span>) مع كل دفع إلى <span dir="ltr">`master`</span> ومع كل طلب سحب (pull request).

## <a id="related-pages"></a>صفحات ذات صلة

- [التثبيت](Installation)
- [إعداد بيئة التطوير](Development-Setup)
- [سجل التغييرات](Changelog)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
