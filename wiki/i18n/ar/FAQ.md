<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# الأسئلة الشائعة حول docxi: مستندات Word في Ruby وRails

إجابات مختصرة عن أكثر الأسئلة تكرارًا، مع روابط لشرح أطول في كل إجابة.

**هل أحتاج إلى تثبيت Microsoft Word أو LibreOffice؟**
لا. تكتب docxi ملفات XML وأرشيف zip بنفسها.

**هل تتطلب docxi إطار Rails؟**
لا. تعمل في أي برنامج Ruby، وتعمل الاختبارات أيضًا مع تحميل Rails من 5.2 إلى 8.1. راجع [التكامل](Integrations).

**كيف أنشئ مستند Word في Ruby on Rails؟**
ابنِه في وحدة تحكم أو في مهمة في الخلفية، ثم أرسل <span dir="ltr">`word.render.string`</span> باستخدام <span dir="ltr">`send_data`</span>. راجع [التكامل](Integrations).

**ما إصدارات Ruby وRails المدعومة؟**
الإصدار 2.5 من Ruby وما بعده. إطار Rails اختياري، وتختبر CI إصدارات Rails من 5.2 إلى 8.1. راجع [التثبيت](Installation).

**هل يمكن لـ docxi تعديل ملف .docx موجود أو ملء قالب Word؟**
لا، فهي تنشئ مستندات جديدة فقط. للقوالب التي تحتوي على حقول دمج، استخدم حزمة قائمة على القوالب مثل sablon.

**هل يُولَّد جدول المحتويات من العناوين؟**
لا. أنت من يضيف كل سطر ورقم صفحته. راجع [الميزات](Features#table-of-contents).

**هل يمكنني استخدام القوائم المرقمة أو أنماط العناوين؟**
ليس في الإصدار 0.1.0. يمكن أن تكون الفقرات عناصر قائمة نقطية (<span dir="ltr">`ul: true`</span>)، ويمكنك تنسيق العناوين بخيارات النص مثل <span dir="ltr">`size:`</span> و<span dir="ltr">`bold:`</span>.

**هل يمكنني تغيير حجم الصفحة أو اتجاهها؟**
ليس بعد. الصفحات بمقاس US Letter وباتجاه عمودي، ولا يمكنك تغيير سوى الهوامش. راجع [الإعدادات](Configuration).

**ما صيغ الصور التي تعمل؟**
صيغة PNG. تُخزَّن كل صورة كجزء PNG.

**هل يمكنني إضافة صورة من عنوان URL؟**
نعم. تقبل <span dir="ltr">`add_media`</span> عناوين URL من نوع <span dir="ltr">`http`</span> و<span dir="ltr">`https`</span> و<span dir="ltr">`ftp`</span>، وتنزّلها عند استدعاء <span dir="ltr">`render`</span>. لا تمرر عناوين URL قادمة من المستخدمين دون التحقق منها. راجع [الأمان](Security).

**هل يمكنني كتابة المستند مباشرة في ملف؟**
تعيد <span dir="ltr">`render`</span> كائن <span dir="ltr">`StringIO`</span>. اكتب قيمة <span dir="ltr">`.string`</span> الخاصة به باستخدام <span dir="ltr">`File.binwrite("out.docx", word.render.string)`</span>.

**هل يمكن لـ docxi بث المستندات الكبيرة جدًا تدريجيًا (streaming)؟**
لا. يُبنى المستند كاملًا في الذاكرة عند استدعاء <span dir="ltr">`render`</span>. راجع [النشر](Deployment).

**هل يمكنني ضبط فقرات من اليمين إلى اليسار؟**
لا توجد خيارات للكتابة من اليمين إلى اليسار في الإصدار 0.1.0، ويُكتب النص كما تمرره.

**أين توجد قائمة المشكلات المعروفة؟**
في [خارطة الطريق](Roadmap)، وفي قسم «Limitations and known issues» (القيود والمشكلات المعروفة) في ملف README.

## <a id="related-pages"></a>صفحات ذات صلة

- [استكشاف الأخطاء وإصلاحها](Troubleshooting)
- [الميزات](Features)
- [البدء](Getting-Started)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
