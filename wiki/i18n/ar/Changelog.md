<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# سجل تغييرات docxi وتاريخ إصداراتها

تلخص هذه الصفحة كل إصدار. الملاحظات الكاملة موجودة في [<span dir="ltr">`CHANGELOG.md`</span>](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) في المستودع، وهو المرجع إذا اختلف الاثنان يومًا.

## <a id="releases"></a>الإصدارات

| الإصدار | التاريخ | ملاحظات |
|---|---|---|
| 0.1.1 | 2026-10-29 | ملخص ووصف أدق للحزمة، ورابط إلى التوثيق. لا تغييرات في الشيفرة |
| 0.1.0 | 2026-10-28 | يعمل مع الإصدارات الحالية من Ruby وRails وrubyzip، ويولّد المستندات نفسها التي يولّدها 0.0.6 |
| <span dir="ltr">0.1.0.rc1</span> | 2026-10-28 | إصدار تجريبي من 0.1.0 |
| 0.0.6 | 2016-10-18 | آخر إصدار قبل تحديث 0.1.0 |
| 0.0.2 | 2016-10-14 | لم تُسجَّل ملاحظات للإصدار |
| 0.0.1 | 2016-10-14 | الإصدار الأول. لم تُسجَّل ملاحظات للإصدار |

لم تُنشر الإصدارات من 0.0.3 إلى 0.0.5 على RubyGems قط.

## <a id="011"></a>الإصدار 0.1.1

لا تغييرات في الشيفرة، والمستندات المولَّدة مطابقة لما يولّده الإصدار 0.1.0.

- أصبح ملخص الحزمة ووصفها يشرحان ما تفعله docxi، بعد أن كانا يصفانها بأنها مولّد قوالب.
- يربط موقع RubyGems هذه الويكي بوصفها توثيق الحزمة.

## <a id="010"></a>الإصدار 0.1.0

**التوافق**

- يتطلب Ruby 2.5 أو أحدث. تبقى التطبيقات التي تعمل بإصدارات Ruby أقدم على 0.0.6.
- يعمل مع Rails أو دونه، ويُختبر إلى جانب Rails من 5.2 إلى 8.1.
- صار قيد nokogiri هو <span dir="ltr">`>= 1.8, < 2`</span> بدلًا من <span dir="ltr">`>= 1.6.3, < 1.7`</span> الذي كان يتعارض مع كل إصدارات Rails منذ 6.0.
- صار قيد rubyzip هو <span dir="ltr">`>= 1.3, < 4`</span> بدلًا من <span dir="ltr">`~> 0.9.9`</span>.
- لم يعد يستدعي <span dir="ltr">`blank?`</span> و<span dir="ltr">`present?`</span> من ActiveSupport، اللتين كانتا تعملان فقط عندما يحمّلهما Rails.

**الإصلاحات**

- لم تعد المستندات المكتوبة باستخدام rubyzip 3 تحتوي على سجلات Zip64، التي يرفضها Google Docs وإصدارات LibreOffice القديمة.
- عاد <span dir="ltr">`add_media`</span> مع عنوان URL لصورة يعمل على Ruby 3.
- لم يعد <span dir="ltr">`add_media`</span> ينفّذ سلسلة تبدأ بالرمز <span dir="ltr">`|`</span> كأمر shell.
- لم تعد <span dir="ltr">`tab`</span> و<span dir="ltr">`frame`</span> تطلقان <span dir="ltr">`NameError`</span>.
- تبني <span dir="ltr">`document.frame("text")`</span> مربع نص يحتوي على النص.
- لم تعد <span dir="ltr">`page_numbers`</span> في رأس الصفحة تطلق <span dir="ltr">`NameError`</span>.

**الحزمة**

- تغيّرت جهة الصيانة، وأُضيفت روابط جديدة للصفحة الرئيسية وسجل التغييرات.
- تتطلب عمليات النشر المصادقة متعددة العوامل (MFA).

## <a id="related-pages"></a>صفحات ذات صلة

- [خارطة الطريق](Roadmap)
- [التثبيت](Installation#upgrading-from-00x)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
