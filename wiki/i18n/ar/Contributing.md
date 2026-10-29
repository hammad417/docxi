<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# المساهمة في docxi

نرحب ببلاغات الأخطاء والإصلاحات والتوثيق والترجمات. تشرح هذه الصفحة كيفية الإبلاغ عن مشكلة وما يحتاج إليه طلب السحب (pull request). لإعداد نسخة عمل، راجع [إعداد بيئة التطوير](Development-Setup).

## <a id="reporting-a-bug"></a>الإبلاغ عن خطأ

افتح بلاغًا (issue) على [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) يتضمن:

- سكربت Ruby قصيرًا يعيد إنتاج المشكلة
- إصدارات Ruby وdocxi وnokogiri وrubyzip (وRails إن كنت تستخدمه)
- ما توقعته وما حصلت عليه. وفي حالة المستندات التالفة، اذكر البرنامج الذي فتحتها به

أبلغ عن المشكلات الأمنية بشكل خاص بدلًا من ذلك. راجع [الأمان](Security).

## <a id="pull-requests"></a>طلبات السحب

1. انسخ المستودع (fork) وأنشئ فرعًا.
2. أجرِ التغيير، والتزم بأسلوب الشيفرة المحيطة.
3. أضف الاختبارات أو حدّثها، وشغّل <span dir="ltr">`bundle exec rake test`</span>.
4. إذا تغيّر XML المولَّد عن قصد، فشغّل <span dir="ltr">`bundle exec rake golden`</span> واشرح التغيير في طلب السحب. تحمي الملفات المرجعية المخرجات التي يعتمد عليها المستخدمون الحاليون.
5. حدّث <span dir="ltr">`CHANGELOG.md`</span> والتوثيق إذا تغيّر السلوك.
6. افتح طلب السحب. تشغّل CI الاختبارات على Ruby من 2.5 إلى 4.0 وRails من 5.2 إلى 8.1، وتتحقق من أن LibreOffice يفتح المستندات النموذجية.

التغييرات التي تعدّل المستندات الحالية، مثل الإصلاحات المخطط لها في الإصدار 0.2.0، مكانها طلب سحب مستقل، كي تسهل مراجعة الفروق في الملفات المرجعية.

## <a id="documentation-and-translations"></a>التوثيق والترجمات

- ملف README الإنجليزي هو مصدر 15 ترجمة في <span dir="ltr">`docs/i18n/`</span>. اتبع <span dir="ltr">`docs/i18n/TRANSLATION_GUIDE.md`</span> في المستودع عند تعديله.
- يُكتب هذا الويكي في مجلد <span dir="ltr">`wiki/`</span> في المستودع ويُنشر منه، ويشرح <span dir="ltr">`wiki/README.md`</span> الطريقة.

## <a id="related-pages"></a>صفحات ذات صلة

- [إعداد بيئة التطوير](Development-Setup)
- [البنية](Architecture)
- [خارطة الطريق](Roadmap)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
