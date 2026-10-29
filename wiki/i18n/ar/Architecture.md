<!-- Translated from wiki/Architecture.md (git blob 38e8491be82a60fc51d016cdfd25fc8316608686). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# بنية docxi: كيف تتحول شيفرة Ruby إلى حزمة .docx

ملف <span dir="ltr">`.docx`</span> أرشيف zip يضم ملفات XML تُسمى أجزاءً (parts)، وتتبع معيار Office Open XML. تحتفظ docxi بكائن Ruby لكل جزء، وتبني XML كل جزء باستخدام nokogiri عند استدعاء <span dir="ltr">`render`</span>، ثم تكتب الأرشيف باستخدام rubyzip. هذه الصفحة موجّهة إلى المساهمين وإلى كل من يحتاج إلى تتبّع أخطاء XML المولَّد.

## <a id="overview"></a>نظرة عامة

```mermaid
flowchart LR
  A["شيفرة Ruby الخاصة بك"] --> B["Docxi::Document"]
  B --> C["Word::Document<br/>الفقرات والجداول والصور<br/>ورؤوس الصفحات وتذييلاتها"]
  C --> D["nokogiri<br/>WordprocessingML XML"]
  D --> E["rubyzip<br/>حزمة .docx"]
  E --> F["StringIO"]
```

## <a id="main-classes"></a>الأصناف الرئيسية

| الصنف | الملف | الدور |
|---|---|---|
| <span dir="ltr">`Docxi::Document`</span> | <span dir="ltr">`lib/docxi/document.rb`</span> | نقطة الدخول. يحتفظ بأجزاء الحزمة وينفّذ <span dir="ltr">`render`</span> |
| <span dir="ltr">`Docxi::Relationships`</span> | <span dir="ltr">`lib/docxi/relationships.rb`</span> | الملف <span dir="ltr">`_rels/.rels`</span>، أي العلاقات على مستوى الحزمة |
| <span dir="ltr">`Docxi::ContentTypes`</span> | <span dir="ltr">`lib/docxi/content_types.rb`</span> | الملف <span dir="ltr">`[Content_Types].xml`</span> |
| <span dir="ltr">`Docxi::Properties`</span> | <span dir="ltr">`lib/docxi/properties.rb`</span> | الملفان <span dir="ltr">`docProps/app.xml`</span> و<span dir="ltr">`docProps/core.xml`</span> |
| <span dir="ltr">`Docxi::Word::Document`</span> | <span dir="ltr">`lib/docxi/word/document.rb`</span> | متن المستند، وإعدادات المقطع، وسجل الصور ورؤوس الصفحات وتذييلاتها والارتباطات |
| <span dir="ltr">`Docxi::Word::Helpers`</span> | <span dir="ltr">`lib/docxi/word/helpers.rb`</span> | الدوال <span dir="ltr">`p`</span> و<span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`tab`</span> و<span dir="ltr">`frame`</span> و<span dir="ltr">`table`</span> و<span dir="ltr">`table_of_content`</span> |
| <span dir="ltr">`Docxi::Word::Contents::*`</span> | <span dir="ltr">`lib/docxi/word/contents/`</span> | الأصناف <span dir="ltr">`Paragraph`</span> و<span dir="ltr">`Text`</span> و<span dir="ltr">`Break`</span> و<span dir="ltr">`Tab`</span> و<span dir="ltr">`Frame`</span> و<span dir="ltr">`Image`</span> و<span dir="ltr">`Table`</span> و<span dir="ltr">`TableOfContent`</span> |
| <span dir="ltr">`Docxi::Word::Headers::Header`</span> و<span dir="ltr">`Docxi::Word::Footers::Footer`</span> | <span dir="ltr">`lib/docxi/word/headers/`</span> و<span dir="ltr">`lib/docxi/word/footers/`</span> | أجزاء رؤوس الصفحات وتذييلاتها، ولكل منها ملف علاقات خاص به |
| <span dir="ltr">`Docxi::Word::Medias`</span> | <span dir="ltr">`lib/docxi/word/medias/`</span> | الصور المضافة عبر <span dir="ltr">`add_media`</span> |
| الأجزاء الثابتة | <span dir="ltr">`lib/docxi/word/*.rb`</span> | الأنماط والترقيم والإعدادات والخطوط والسمة والتأثيرات والحواشي السفلية والتعليقات الختامية |

## <a id="what-render-does"></a>ما الذي تفعله <span dir="ltr">`render`</span>

1. تفتح <span dir="ltr">`Docxi::Document#render`</span> الكائن <span dir="ltr">`Zip::OutputStream.write_buffer`</span>.
2. تكتب <span dir="ltr">`Word::Document#render`</span> الملف <span dir="ltr">`word/document.xml`</span> (يكتب كل عنصر XML الخاص به)، ثم الأجزاء الثابتة ورؤوس الصفحات وتذييلاتها والصور و<span dir="ltr">`word/_rels/document.xml.rels`</span>.
3. تُكتب خصائص الحزمة وعلاقات الحزمة وأنواع المحتوى في النهاية.
4. يُعاد المخزن المؤقت ككائن <span dir="ltr">`StringIO`</span>.

يمر كل جزء عبر <span dir="ltr">`Docxi.write_entry`</span>، التي تضبط حجم المُدخل قبل كتابته. ولولا ذلك لأضافت rubyzip 3 سجلات Zip64 إلى كل مُدخل، في حين يرفض Google Docs وإصدارات LibreOffice القديمة فتح مثل هذه الملفات.

## <a id="parts-in-a-generated-document"></a>الأجزاء في المستند المولَّد

| الجزء | المحتوى |
|---|---|
| <span dir="ltr">`[Content_Types].xml`</span> | أنواع محتوى الأجزاء |
| <span dir="ltr">`_rels/.rels`</span> | روابط إلى المستند وخصائصه |
| <span dir="ltr">`docProps/app.xml`</span> و<span dir="ltr">`docProps/core.xml`</span> | اسم التطبيق («Docxi Ruby Gem») ووقت التوليد |
| <span dir="ltr">`word/document.xml`</span> | المتن: الفقرات والجداول وإعدادات المقطع (الهوامش وحجم الصفحة ومراجع رؤوس الصفحات وتذييلاتها) |
| <span dir="ltr">`word/styles.xml`</span> و<span dir="ltr">`word/stylesWithEffects.xml`</span> | الأنماط. الخط الافتراضي Arial بحجم 12 نقطة |
| <span dir="ltr">`word/numbering.xml`</span> | تعريف القائمة النقطية |
| <span dir="ltr">`word/settings.xml`</span> و<span dir="ltr">`word/webSettings.xml`</span> و<span dir="ltr">`word/fontTable.xml`</span> | إعدادات المستند والخطوط |
| <span dir="ltr">`word/theme/theme1.xml`</span> | السمة |
| <span dir="ltr">`word/footnotes.xml`</span> و<span dir="ltr">`word/endnotes.xml`</span> | فواصل فقط. لا تملك docxi واجهة برمجية للحواشي السفلية |
| <span dir="ltr">`word/headerN.xml`</span> و<span dir="ltr">`word/footerN.xml`</span> | جزء لكل رأس صفحة وتذييل، مرقّمة بترتيب إضافتها |
| <span dir="ltr">`word/media/imageN.png`</span> | جزء لكل استدعاء لـ <span dir="ltr">`add_media`</span> |
| <span dir="ltr">`word/_rels/document.xml.rels`</span> | علاقات المتن |

لمعاينة محتوى مستند، فك ضغطه، مثلًا باستخدام <span dir="ltr">`unzip -p report.docx word/document.xml`</span>.

## <a id="relationship-ids"></a>معرّفات العلاقات

يبدأ <span dir="ltr">`word/_rels/document.xml.rels`</span> بتسع علاقات ثابتة، من <span dir="ltr">`rId1`</span> إلى <span dir="ltr">`rId9`</span>:

1. الأنماط
2. الأنماط مع التأثيرات
3. الإعدادات
4. إعدادات الويب
5. جدول الخطوط
6. السمة
7. التعليقات الختامية
8. الحواشي السفلية
9. الترقيم

تليها الصور ورؤوس الصفحات وتذييلاتها والارتباطات التشعبية، مرقّمةً بالترتيب الذي تستدعي به <span dir="ltr">`add_media`</span> و<span dir="ltr">`add_header`</span> و<span dir="ltr">`add_footer`</span> و<span dir="ltr">`add_hyperlink`</span>. تعيد <span dir="ltr">`add_hyperlink`</span> معرّفها كي تتمكن <span dir="ltr">`hyperlink`</span> من الإشارة إليه.

## <a id="units-and-conversions"></a>الوحدات والتحويلات

- يُضاعَف <span dir="ltr">`size`</span> للنص، لأن Word يخزّن أنصاف النقاط.
- يُضرب <span dir="ltr">`columns_width`</span> و<span dir="ltr">`width`</span> للخلية في 14.92 للحصول على أجزاء من عشرين من النقطة.
- تحسب الصور المحددة الموضع وصور رؤوس الصفحات وصور خلايا الجداول حجمها على أنه <span dir="ltr">`width * height * 14.92`</span> EMU. هذه مشكلة معروفة مخطط إصلاحها في الإصدار 0.2.0.
- تُمرَّر الهوامش والمسافات البادئة وارتفاعات الصفوف كما هي، بوحدة جزء من عشرين من النقطة.

## <a id="how-output-stays-stable"></a>كيف تبقى المخرجات ثابتة

يحتوي <span dir="ltr">`test/golden`</span> على كل جزء ولّده الإصدار 0.0.6 لمستندين نموذجيين (<span dir="ltr">`test/support/sample_documents.rb`</span>). تولّد الاختبارات المستندات نفسها وتقارن كل جزء بصيغة XML قانونية (canonical)، فلا تُحتسب الفروق بين إصدارات nokogiri، بينما يُكتشف أي تغيير حقيقي. يعيد <span dir="ltr">`bundle exec rake golden`</span> كتابة الملفات المرجعية (golden files) بعد أي تغيير مقصود. راجع [إعداد بيئة التطوير](Development-Setup).

## <a id="related-pages"></a>صفحات ذات صلة

- [مرجع API](API-Reference)
- [الميزات](Features)
- [المساهمة](Contributing)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
