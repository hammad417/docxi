<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# إعداد docxi: هوامش الصفحة وإعداد الصفحة والوحدات

ليس لمكتبة docxi ملف إعدادات ولا متغيرات بيئة ولا إعدادات عامة، فكل شيء يُضبط لكل مستند أو لكل عنصر على حدة. تشرح هذه الصفحة الإعدادات على مستوى المستند والوحدات التي تستخدمها الخيارات. أما خيارات كل عنصر فمدرجة في صفحة [الميزات](Features).

## <a id="page-margins"></a>هوامش الصفحة

تُمرَّر الهوامش عند إنشاء المستند:

```ruby
require "docxi"

# جميع القيم بوحدة جزء من عشرين من النقطة، و1440 تساوي بوصة واحدة
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| الخيار | القيمة الافتراضية | المعنى |
|---|---|---|
| <span dir="ltr">`top`</span> | <span dir="ltr">`720`</span> ‏(0.5 بوصة) | الهامش العلوي |
| <span dir="ltr">`right`</span> | <span dir="ltr">`1440`</span> ‏(بوصة واحدة) | الهامش الأيمن |
| <span dir="ltr">`bottom`</span> | <span dir="ltr">`720`</span> ‏(0.5 بوصة) | الهامش السفلي |
| <span dir="ltr">`left`</span> | <span dir="ltr">`1440`</span> ‏(بوصة واحدة) | الهامش الأيسر |
| <span dir="ltr">`header`</span> | <span dir="ltr">`190`</span> | المسافة من الحافة العليا للصفحة إلى رأس الصفحة |
| <span dir="ltr">`footer`</span> | <span dir="ltr">`200`</span> | المسافة من الحافة السفلى للصفحة إلى تذييل الصفحة |

القاموس (Hash) الذي تمرره يحل محل القيم الافتراضية، لذا اضبط <span dir="ltr">`top`</span> و<span dir="ltr">`right`</span> و<span dir="ltr">`bottom`</span> و<span dir="ltr">`left`</span> معًا. أي واحد من هذه الأربعة لا تضبطه يُكتب قيمةً فارغة. أما <span dir="ltr">`header`</span> و<span dir="ltr">`footer`</span> فيعودان إلى 190 و200 عندما يكونان مفقودين أو فارغين.

## <a id="page-size-and-orientation"></a>حجم الصفحة واتجاهها

الصفحات بمقاس US Letter وباتجاه عمودي: 12240 × 15840 جزءًا من عشرين من النقطة، أي 8.5 × 11 بوصة. لا يمكن تغيير الحجم ولا الاتجاه في الإصدار 0.1.0.

## <a id="first-page"></a>الصفحة الأولى

تفعّل docxi دائمًا خيار «صفحة أولى مختلفة» في Word. تُستخدم رؤوس الصفحات وتذييلاتها المضافة عبر <span dir="ltr">`add_header`</span> و<span dir="ltr">`add_footer`</span> بدءًا من الصفحة 2. أما الصفحة 1 فتستخدم <span dir="ltr">`add_first_header`</span> و<span dir="ltr">`add_first_footer`</span>، لذا اضبطهما أيضًا إذا احتاجت الصفحة الأولى إلى رأس أو تذييل. راجع [الميزات](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>ما يحتويه كل مستند

- الخط الافتراضي هو Arial بحجم 12 نقطة.
- الأنماط المعرّفة: <span dir="ltr">`Normal`</span> و<span dir="ltr">`DefaultParagraphFont`</span> و<span dir="ltr">`TableNormal`</span> و<span dir="ltr">`TableCustom`</span> و<span dir="ltr">`NoList`</span>.
- تستخدم عناصر القائمة النقطية (<span dir="ltr">`ul: true`</span>) شرطةً كعلامة نقطية.
- تذكر خصائص المستند أن اسم التطبيق «Docxi Ruby Gem»، وتسجّل وقت توليد المستند.

## <a id="units"></a>الوحدات

| الخيار | الوحدة |
|---|---|
| الخيار <span dir="ltr">`size:`</span> للنص | النقاط (تضاعفها docxi لتحويلها إلى أنصاف النقاط التي يستخدمها Word) |
| الهوامش، و<span dir="ltr">`left:`</span> و<span dir="ltr">`right:`</span> للفقرة، و<span dir="ltr">`iwidth:`</span> للجدول، و<span dir="ltr">`height:`</span> للصف | جزء من عشرين من النقطة (1440 لكل بوصة) |
| أحجام <span dir="ltr">`borders:`</span> للخلية | ثُمن النقطة |
| الخيار <span dir="ltr">`columns_width:`</span> و<span dir="ltr">`width:`</span> للخلية | تُضرب في 14.92 للحصول على أجزاء من عشرين من النقطة، وهي قريبة من البكسلات عند 96 نقطة في البوصة |
| الخيار <span dir="ltr">`style:`</span> للصورة المضمّنة | أطوال CSS مثل <span dir="ltr">`"120pt"`</span> |
| الخياران <span dir="ltr">`pH:`</span> و<span dir="ltr">`pV:`</span> للصورة المحددة الموضع | وحدة EMU ‏(914400 لكل بوصة) |
| ألوان النص والتعبئة | قيمة ست عشرية دون <span dir="ltr">`#`</span>، مثل <span dir="ltr">`"8B0000"`</span> |
| الخيار <span dir="ltr">`fill_color:`</span> للصورة | لون VML مع <span dir="ltr">`#`</span>، مثل <span dir="ltr">`"#00FF00"`</span> |

## <a id="related-pages"></a>صفحات ذات صلة

- [الميزات](Features)
- [مرجع API](API-Reference)
- [أفضل الممارسات](Best-Practices)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
