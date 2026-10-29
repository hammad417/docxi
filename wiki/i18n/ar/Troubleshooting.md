<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# استكشاف أخطاء docxi وإصلاحها: أخطاء التثبيت وأخطاء التشغيل والمخرجات التالفة

ابحث عن العَرَض الذي تواجهه أدناه. معظم مشكلات التطبيقات القديمة سببها docxi 0.0.x، والترقية إلى الإصدار 0.1.0 تحلها دون تغيير الشيفرة.

## <a id="installation"></a>التثبيت

**لا يجد Bundler إصدارات متوافقة من <span dir="ltr">`nokogiri`</span> في تطبيق Rails 6 أو أحدث.**
يقيّد docxi 0.0.6 الاعتمادية بـ <span dir="ltr">`nokogiri < 1.7`</span>، بينما تحتاج كل إصدارات Rails منذ 6.0 إلى nokogiri 1.8.5 أو أحدث. استخدم <span dir="ltr">`gem "docxi", "~> 0.1"`</span>.

**لا يجد Bundler إصدارات متوافقة من <span dir="ltr">`rubyzip`</span>.**
يقيّد docxi 0.0.6 الاعتمادية بـ <span dir="ltr">`rubyzip ~> 0.9.9`</span>، وهذا يتعارض مع أي حزمة تحتاج إلى إصدار أحدث من rubyzip. أما docxi 0.1.0 فتقبل rubyzip من 1.3 إلى 3.x.

**يستمر Bundler في تثبيت الإصدار 0.0.6.**
السبب إمّا أن ملف Gemfile يثبّت docxi على إصدار 0.0.x، مثل <span dir="ltr">`"~> 0.0.6"`</span>، وإمّا أن إصدار Ruby أقدم من 2.5 الذي يتطلبه الإصدار 0.1.0. غيّر السطر في Gemfile إلى <span dir="ltr">`gem "docxi", "~> 0.1"`</span> وتحقق من <span dir="ltr">`ruby -v`</span>.

**فشل تصريف nokogiri.**
لا توجد حزمة nokogiri مُصرَّفة مسبقًا تطابق منصتك، لذا تُبنى من الشيفرة المصدرية. ثبّت مصرّف C واتبع [دليل تثبيت nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>أخطاء أثناء بناء المستند

**ظهور <span dir="ltr">`NoMethodError: undefined method 'present?'`</span> خارج Rails.**
كان docxi 0.0.6 يفترض أن ActiveSupport محمّلة. أما 0.1.0 فلا تحتاج إليها.

**ظهور <span dir="ltr">`NameError`</span> عند استدعاء <span dir="ltr">`tab`</span> أو <span dir="ltr">`frame`</span> أو <span dir="ltr">`page_numbers`</span> في رأس الصفحة.**
كانت هذه الدوال معطّلة في 0.0.6 وتعمل في 0.1.0، بما في ذلك <span dir="ltr">`document.frame("text")`</span>.

**ظهور <span dir="ltr">`Errno::ENOENT`</span> عند إضافة صورة.**
المسار غير موجود بالنسبة إلى المجلد الذي يعمل فيه برنامجك. على Ruby 3 مع docxi 0.0.6 تفشل عناوين URL للصور بالطريقة نفسها أيضًا، لأن 0.0.6 كان يستخدم <span dir="ltr">`Kernel#open`</span>. أما 0.1.0 فتنزّل عناوين URL باستخدام <span dir="ltr">`URI#open`</span>.

## <a id="documents-that-dont-look-right"></a>مستندات لا تظهر كما ينبغي

**لا يفتح Google Docs أو LibreOffice الملف.**
يكتب docxi 0.0.6 مع rubyzip 3 سجلات Zip64 في كل مُدخل، وهذه البرامج ترفضها. أما 0.1.0 فتكتب الأرشيفات دونها.

**لا يوجد رأس أو تذييل في الصفحة الأولى.**
تمنح docxi الصفحة الأولى رأسًا وتذييلًا خاصين بها. اضبطهما باستخدام <span dir="ltr">`add_first_header`</span> و<span dir="ltr">`add_first_footer`</span>. راجع [الإعدادات](Configuration#first-page).

**أحد الهوامش مفقود، أو يبدو تخطيط الصفحة خاطئًا.**
عند ضبط الهوامش، مرّر <span dir="ltr">`top`</span> و<span dir="ltr">`right`</span> و<span dir="ltr">`bottom`</span> و<span dir="ltr">`left`</span> معًا. ما تتركه منها يُكتب فارغًا.

**تختفي المسافات بين المقاطع النصية.**
يحذف Word المسافات في بداية المقطع أو نهايته. أضف <span dir="ltr">`space: true`</span> إلى استدعاء <span dir="ltr">`text`</span> ذاك.

**الصور في رؤوس الصفحات أو في خلايا الجداول أو التي تستخدم <span dir="ltr">`pH`</span>/<span dir="ltr">`pV`</span> صغيرة جدًا أو مربعة.**
مشكلة معروفة في 0.1.0: يُحسب حجمها على أنه <span dir="ltr">`width * height * 14.92`</span> EMU في البعدين كليهما. استخدم صورًا مضمّنة يُضبط حجمها عبر <span dir="ltr">`style:`</span>، أو راجع الحل البديل في [أفضل الممارسات](Best-Practices#images).

**تختفي صورة عند توليد المستند نفسه للمرة الثانية.**
الصور المضافة من كائن <span dir="ltr">`File`</span> مفتوح أو من أي IO آخر تُقرأ أثناء <span dir="ltr">`render`</span>، والتوليد الثاني يقرأ دفقًا فارغًا. مرّر مسارًا بدلًا من ذلك، أو ولّد المستند مرة واحدة.

**تظهر الحدود المضبوطة على <span dir="ltr">`table`</span> حول كل خلية.**
مشكلة معروفة: يُطبَّق الخيار <span dir="ltr">`borders:`</span> على مستوى الجدول على كل خلية بدلًا من الحواف الخارجية للجدول. ولرسم الإطار الخارجي فقط، اضبط <span dir="ltr">`borders:`</span> على الخلايا الواقعة على أطراف الجدول، لجوانبها الخارجية فقط.

**تبدأ أرقام الصفحات في خلايا الجداول بالنص «GlobalOptions».**
مشكلة معروفة في <span dir="ltr">`TableCell#page_numbers`</span>. ضع أرقام الصفحات في تذييل أو رأس صفحة بدلًا من ذلك.

**صورة JPEG أو GIF تظهر بشكل خاطئ أو لا تظهر.**
تُخزَّن كل صورة كجزء PNG بنوع محتوى PNG أيًا كانت صيغتها الحقيقية، لذا قد لا تعرض بعض البرامج الصيغ الأخرى. حوّل الصور إلى PNG قبل إضافتها.

## <a id="inspecting-a-generated-file"></a>فحص ملف مولَّد

ملف <span dir="ltr">`.docx`</span> أرشيف zip، لذا يمكنك النظر داخله:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

إذا استمرت المشكلة، [افتح بلاغًا (issue)](https://github.com/hammad417/docxi/issues) يتضمن سكربتًا قصيرًا يعيد إنتاجها، وإصدارات Ruby وdocxi وnokogiri وrubyzip.

## <a id="related-pages"></a>صفحات ذات صلة

- [الأسئلة الشائعة](FAQ)
- [أفضل الممارسات](Best-Practices)
- [التثبيت](Installation)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
