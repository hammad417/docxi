<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# أفضل الممارسات مع docxi: تجنّب المفاجآت المعتادة

هذه العادات مستمدة مباشرة من طريقة عمل docxi، ومعظمها يمنع مشكلة مدرجة في [استكشاف الأخطاء وإصلاحها](Troubleshooting). وهي مجمّعة حسب الموضوع. نصائح إعداد الصفحة تنطبق على الجميع، أما نصائح الأمان والأداء فهي الأهم لتطبيقات الإنتاج.

## <a id="page-setup"></a>إعداد الصفحة

- **مرّر الهوامش الأربعة معًا.** القاموس الذي تمرره إلى <span dir="ltr">`Docxi::Document.new`</span> يحل محل القيم الافتراضية، وأي واحد من <span dir="ltr">`top`</span> و<span dir="ltr">`right`</span> و<span dir="ltr">`bottom`</span> و<span dir="ltr">`left`</span> تتركه يُكتب فارغًا.
- **اضبط رأس الصفحة الأولى وتذييلها.** تفعّل docxi دائمًا صفحة أولى منفصلة. إذا احتاجت الصفحة 1 إلى رأس أو تذييل، فاستدعِ <span dir="ltr">`add_first_header`</span> و<span dir="ltr">`add_first_footer`</span> أيضًا.

## <a id="text"></a>النص

- **استخدم <span dir="ltr">`space: true`</span> للمسافات في البداية أو النهاية.** وإلا حذفها Word والتصقت كلمات المقطعين.
- **ضع التنسيق المشترك على الفقرة.** الخيارات المُمرَّرة إلى <span dir="ltr">`p`</span>، مثل <span dir="ltr">`size:`</span> أو <span dir="ltr">`font:`</span>، تنطبق على كل مقطع بداخلها، فيقل ما تحتاج إلى تكراره.
- **استخدم <span dir="ltr">`p.br`</span> لفواصل الأسطر و<span dir="ltr">`document.br(page: true)`</span> لفواصل الصفحات.** استدعاء <span dir="ltr">`document.br`</span> دون <span dir="ltr">`page: true`</span> ينتج XML غير صالح.

## <a id="tables"></a>الجداول

- **اضبط الحدود على الصفوف أو الخلايا، لا على الجدول.** في الإصدار 0.1.0 يُطبَّق الخيار <span dir="ltr">`borders:`</span> على مستوى الجدول على كل خلية بدلًا من الحواف الخارجية للجدول.
- **حدّد <span dir="ltr">`columns_width`</span> لكل عمود.** تأخذ الخلايا عرضها منه، فتحافظ الجداول على التخطيط الذي تتوقعه.

## <a id="images"></a>الصور

- **استخدم ملفات PNG.** تُخزَّن كل صورة كجزء PNG.
- **اضبط حجم الصور المضمّنة باستخدام <span dir="ltr">`style:`</span>**، مثل <span dir="ltr">`style: { width: "120pt", height: "60pt" }`</span>. إنه حجم الصور الوحيد الذي يعمل كما تتوقع في الإصدار 0.1.0.
- **مرّر مسارًا بدلًا من كائن <span dir="ltr">`File`</span> مفتوح إذا كان المستند قد يُولَّد أكثر من مرة.** صور IO تُقرأ أثناء <span dir="ltr">`render`</span>، والتوليد الثاني يحصل على صورة فارغة.
- **صور رؤوس الصفحات وخلايا الجداول والصور المحددة الموضع حتى الإصدار 0.2.0.** يُحسب عرضها وارتفاعها كلاهما على أنه <span dir="ltr">`width * height * 14.92`</span> EMU، لذا تظهر مربعة. للحصول على صورة مربعة طول ضلعها <span dir="ltr">`S`</span> بوصة، مرّر <span dir="ltr">`width`</span> و<span dir="ltr">`height`</span> بقيمة <span dir="ltr">`Math.sqrt(S * 914400 / 14.92).round`</span>، وهي نحو 248 لبوصة واحدة.

## <a id="security"></a>الأمان

- **لا تمرر سلاسل نصية غير موثوقة إلى <span dir="ltr">`add_media`</span>.** المسار يجعل docxi تقرأ ذلك الملف من خادمك، وعنوان URL يجعل خادمك ينزّله. مرّر كائن IO فتحته بنفسك، أو تحقق من القيمة أولًا. راجع [الأمان](Security).
- **يُهرَّب النص تلقائيًا.** السلاسل المُمرَّرة إلى <span dir="ltr">`text`</span> تُكتب كنص XML، لذا لا يمكن لمدخلات المستخدمين تغيير بنية المستند.

## <a id="performance-and-reliability"></a>الأداء والموثوقية

- **ابنِ المستندات الكبيرة في مهمة في الخلفية.** يُبنى المستند كاملًا في الذاكرة عند استدعاء <span dir="ltr">`render`</span>. راجع [التكامل](Integrations).
- **ولّد مرة واحدة ثم أعد استخدام البايتات.** احتفظ بـ <span dir="ltr">`word.render.string`</span> إذا احتجت إلى إرسال المستند أو تخزينه أكثر من مرة.
- **ثبّت إصدارًا متوافقًا** باستخدام <span dir="ltr">`gem "docxi", "~> 0.1"`</span>.
- **اختبر ما تولّده.** ملف <span dir="ltr">`.docx`</span> أرشيف zip من ملفات XML، لذا يمكن للاختبارات فك ضغطه وفحص XML. اختبارات docxi نفسها تعمل بهذه الطريقة. راجع [البنية](Architecture).

## <a id="related-pages"></a>صفحات ذات صلة

- [الميزات](Features)
- [الإعدادات](Configuration)
- [استكشاف الأخطاء وإصلاحها](Troubleshooting)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
