<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# ويكي docxi: إنشاء مستندات Word ‏(.docx) باستخدام Ruby

مكتبة docxi حزمة (gem) للغة Ruby تُنشئ ملفات Microsoft Word بصيغة <span dir="ltr">`.docx`</span> من خلال الشيفرة البرمجية. تدعم الفقرات والنص المنسّق والجداول والصور ورؤوس الصفحات وتذييلاتها وأرقام الصفحات والارتباطات التشعبية وجدول محتويات بسيطًا. لا تحتاج إلى Microsoft Word أو LibreOffice أو قالب، وتعمل في برامج Ruby العادية كما تعمل في تطبيقات Ruby on Rails.

يقدّم ملف [README](https://github.com/hammad417/docxi#readme) تعريفًا مختصرًا، أما هذا الويكي فيغطي كل الخيارات، ويضم وصفات جاهزة للنسخ، ويشرح طريقة عمل docxi من الداخل، ويساعدك عندما يحدث خطأ ما.

الإصدار الحالي: **0.1.1**، ويتطلب Ruby 2.5 أو أحدث. راجع [سجل التغييرات](Changelog).

## <a id="start-here"></a>ابدأ من هنا

**البداية مع docxi**

- [البدء](Getting-Started): ثبّت الحزمة وأنشئ مستندك الأول
- [التثبيت](Installation): المتطلبات، وBundler، والترقية من 0.0.x
- [أمثلة الاستخدام](Usage-Examples): تقارير وفواتير وخطابات وجداول مبنية من البيانات، جاهزة للنسخ

**بناء المستندات**

- [الميزات](Features): كل عنصر والخيارات التي يقبلها
- [الإعدادات](Configuration): الهوامش وإعداد الصفحة والوحدات
- [مرجع API](API-Reference): الأصناف والدوال في لمحة
- [أفضل الممارسات](Best-Practices): عادات تجنّبك المفاجآت المعتادة

**عندما يحدث خطأ ما**

- [استكشاف الأخطاء وإصلاحها](Troubleshooting): رسائل الخطأ والمستندات التي لا تظهر كما ينبغي
- [الأسئلة الشائعة](FAQ): إجابات مختصرة عن الأسئلة الشائعة
- [الأمان](Security): الإبلاغ عن المشكلات والتعامل مع المدخلات غير الموثوقة

**للتعمق أكثر**

- [البنية](Architecture): كيف تتحول استدعاءات Ruby إلى حزمة <span dir="ltr">`.docx`</span>
- [التكامل](Integrations): وحدات التحكم في Rails وتطبيقات Rack والمهام في الخلفية
- [النشر](Deployment): تشغيل docxi في بيئة الإنتاج، وطريقة نشر الإصدارات
- [المساهمة](Contributing) و[إعداد بيئة التطوير](Development-Setup)
- [خارطة الطريق](Roadmap) و[سجل التغييرات](Changelog) و[الشكر والتقدير](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>ما الذي تفعله docxi وما الذي لا تفعله

تُنشئ docxi مستندات جديدة. لا يمكنها فتح ملف <span dir="ltr">`.docx`</span> موجود أو تعديله، ولا ملء قالب Word، ولا تنفيذ دمج المراسلات. ليس فيها قوائم مرقمة ولا أنماط عناوين ولا جدول محتويات تلقائي، والصفحات دائمًا بمقاس US Letter وباتجاه عمودي. المشكلات المعروفة المخطط إصلاحها في الإصدار 0.2.0 مدرجة في [خارطة الطريق](Roadmap).

## <a id="credits"></a>الشكر والتقدير

صيانة docxi مسؤولية [Hammad Maqbool](https://github.com/hammad417) في [Phaedra Solutions](https://www.phaedrasolutions.com/). بدأت المكتبة عام 2016 انطلاقًا من [docxer](https://github.com/gtimti/docxer) من تأليف Timur Gabdrakipov. راجع صفحة [الشكر والتقدير](Credits) لمزيد من التفاصيل.

## <a id="languages"></a>اللغات

النسخة الإنجليزية هي النسخة المرجعية لهذا الويكي. روابط الترجمات موجودة في الشريط الجانبي وفي أعلى كل صفحة.

---

[الإبلاغ عن مشكلة](https://github.com/hammad417/docxi/issues) · [README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi)

</div>
