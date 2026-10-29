<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# ميزات docxi: الفقرات والنصوص والجداول والصور ورؤوس الصفحات وتذييلاتها

تسرد هذه الصفحة كل عنصر يمكن لمكتبة docxi إضافته إلى مستند Word، والخيارات التي يقبلها كل عنصر. تعتمد الأمثلة على الإعداد التالي:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

يمكنك حذف الخيارات التي لا تحتاج إليها. الألوان قيم ست عشرية دون <span dir="ltr">`#`</span>، مثل <span dir="ltr">`"8B0000"`</span>. للاطلاع على الوحدات وإعداد الصفحة، راجع [الإعدادات](Configuration).

## <a id="paragraphs"></a>الفقرات

يضيف <span dir="ltr">`document.p(options) { |p| ... }`</span> فقرة.

| الخيار | القيم | التأثير |
|---|---|---|
| <span dir="ltr">`align`</span> | <span dir="ltr">`"left"`</span> و<span dir="ltr">`"center"`</span> و<span dir="ltr">`"right"`</span> و<span dir="ltr">`"both"`</span> | المحاذاة. القيمة <span dir="ltr">`"both"`</span> تضبط النص من الجانبين |
| <span dir="ltr">`fill`</span> | لون ست عشري | تظليل خلف الفقرة ونصها |
| <span dir="ltr">`left`</span> و<span dir="ltr">`right`</span> | جزء من عشرين من النقطة | المسافة البادئة. اضبط الاثنين معًا، لأن ضبط أحدهما وحده يُتجاهل |
| <span dir="ltr">`ul`</span> | <span dir="ltr">`true`</span> | يجعل الفقرة عنصرًا في قائمة نقطية |
| <span dir="ltr">`bottom`</span> | <span dir="ltr">`true`</span> | يرسم حدًا أسفل الفقرة |

يمكن أيضًا تمرير خيارات النص مثل <span dir="ltr">`size:`</span> أو <span dir="ltr">`bold:`</span> أو <span dir="ltr">`color:`</span> إلى <span dir="ltr">`p`</span>، فتصبح القيم الافتراضية لكل مقطع نصي في الفقرة.

داخل الكتلة يمكنك استدعاء <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`tab`</span> و<span dir="ltr">`image`</span> و<span dir="ltr">`hyperlink`</span> و<span dir="ltr">`frame`</span>. والدالة <span dir="ltr">`document.text(string, options)`</span> اختصار لفقرة تحتوي على مقطع واحد.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>النص

يضيف <span dir="ltr">`p.text(string, options)`</span> مقطعًا نصيًا (run).

| الخيار | القيم | التأثير |
|---|---|---|
| <span dir="ltr">`bold`</span> | <span dir="ltr">`true`</span> | غامق |
| <span dir="ltr">`italic`</span> | <span dir="ltr">`true`</span> | مائل |
| <span dir="ltr">`underline`</span> | <span dir="ltr">`true`</span> | تسطير مفرد |
| <span dir="ltr">`color`</span> | لون ست عشري | لون النص |
| <span dir="ltr">`size`</span> | نقاط | حجم الخط |
| <span dir="ltr">`font`</span> | اسم الخط | الخط، مثل <span dir="ltr">`"Georgia"`</span> |
| <span dir="ltr">`fill`</span> | لون ست عشري | تظليل خلف النص |
| <span dir="ltr">`space`</span> | <span dir="ltr">`true`</span> | يحافظ على المسافات في بداية النص ونهايته |
| <span dir="ltr">`br`</span> | <span dir="ltr">`true`</span> | يضيف فاصل أسطر بعد النص |

تُتجاهل السلسلة النصية <span dir="ltr">`nil`</span>.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>فواصل الأسطر وفواصل الصفحات وعلامات الجدولة

- يضيف <span dir="ltr">`p.br`</span> أو <span dir="ltr">`p.br(times: 3)`</span> فواصل أسطر داخل الفقرة.
- يبدأ <span dir="ltr">`document.br(page: true)`</span> صفحة جديدة.
- يضيف <span dir="ltr">`p.tab`</span> أو <span dir="ltr">`p.tab(times: 2)`</span> علامات جدولة داخل الفقرة.
- يضيف <span dir="ltr">`document.tab(page: true)`</span> فقرة تحتوي على علامة جدولة واحدة.

لا تستدعِ <span dir="ltr">`document.br`</span> أو <span dir="ltr">`document.tab`</span> دون <span dir="ltr">`page: true`</span>، فهما تضعان مقطعًا نصيًا مباشرة في متن المستند، وهذا ليس WordprocessingML صالحًا.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>الجداول

تقبل <span dir="ltr">`document.table(options)`</span> كتلة، وداخلها تضيف <span dir="ltr">`table.tr(options)`</span> الصفوف وتضيف <span dir="ltr">`tr.tc(options)`</span> الخلايا.

**خيارات الجدول**

| الخيار | القيم | التأثير |
|---|---|---|
| <span dir="ltr">`columns_width`</span> | مصفوفة أرقام | عرض الأعمدة، مضروبًا في 14.92 لتحويله إلى أجزاء من عشرين من النقطة. وهو أيضًا العرض الافتراضي للخلايا |
| <span dir="ltr">`height`</span> | جزء من عشرين من النقطة | الحد الأدنى لارتفاع كل صف |
| <span dir="ltr">`iwidth`</span> | جزء من عشرين من النقطة | إزاحة الجدول عن الهامش الأيسر |
| <span dir="ltr">`style`</span> | معرّف النمط | نمط الجدول. تعرّف docxi النمطين <span dir="ltr">`TableNormal`</span> و<span dir="ltr">`TableCustom`</span> |
| <span dir="ltr">`width`</span> | رقم | العرض المفضّل للجدول، ويُكتب بالنوع <span dir="ltr">`auto`</span>. ومن دون <span dir="ltr">`columns_width`</span> تحصل كل خلية أيضًا على هذا العرض مضروبًا في 14.92 |
| <span dir="ltr">`align`</span> | <span dir="ltr">`"left"`</span> و<span dir="ltr">`"center"`</span> و<span dir="ltr">`"right"`</span> | محاذاة الجدول. ويُحاذى النص المضاف باستخدام <span dir="ltr">`tc.text`</span> بالطريقة نفسها |
| <span dir="ltr">`padding`</span> | <span dir="ltr">`true`</span> | يضيف حشوة ثابتة داخل الخلايا |
| <span dir="ltr">`borders`</span> | قاموس (Hash) | مشكلة معروفة: يُطبَّق على كل خلية بدلًا من الحواف الخارجية للجدول |

**خيارات الصف**: يضبط <span dir="ltr">`height:`</span> الحد الأدنى لارتفاع الصف. وأي خيار خلية يُمرَّر إلى <span dir="ltr">`tr`</span> ينطبق على كل خلايا الصف. وتُمرَّر خيارات الجدول بالطريقة نفسها إلى كل صف وكل خلية.

**خيارات الخلية**

| الخيار | القيم | التأثير |
|---|---|---|
| <span dir="ltr">`borders`</span> | <span dir="ltr">`{ top:, right:, bottom:, left: }`</span> | عرض الحد بوحدة ثُمن النقطة. القيمة <span dir="ltr">`nil`</span> تعني عدم وجود حد في ذلك الجانب |
| <span dir="ltr">`fill`</span> | لون ست عشري | تظليل الخلية. الخلايا التي لا تعبئة لها بيضاء (<span dir="ltr">`"FFFFFF"`</span>) |
| <span dir="ltr">`valign`</span> | <span dir="ltr">`"top"`</span> و<span dir="ltr">`"center"`</span> و<span dir="ltr">`"bottom"`</span> | المحاذاة العمودية |
| <span dir="ltr">`colspan`</span> | عدد صحيح | عدد الأعمدة التي تمتد عليها الخلية |
| <span dir="ltr">`rowspan`</span> | <span dir="ltr">`true`</span> | يبدأ دمجًا عموديًا |
| <span dir="ltr">`merged`</span> | <span dir="ltr">`true`</span> | يكمل الدمج العمودي من الخلية التي فوقها، ويُتجاهل محتواها |
| <span dir="ltr">`width`</span> | رقم | عرض الخلية، مضروبًا في 14.92 |

**محتوى الخلية**

- الدالة <span dir="ltr">`tc.text(string, options)`</span>: تنطبق خيارات الخلية والصف، مثل <span dir="ltr">`bold:`</span> أو <span dir="ltr">`align:`</span>، على النص أيضًا.
- الدالة <span dir="ltr">`tc.p(options) { |p| ... }`</span>: فقرة بخياراتها الخاصة فقط.
- الدالة <span dir="ltr">`tc.image(media, width:, height:)`</span>: صورة. راجع مشكلة الحجم في قسم [الصور](#images).
- الدالة <span dir="ltr">`tc.page_numbers`</span>: حقل رقم الصفحة. مشكلة معروفة: يسبقه النص «GlobalOptions».

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>الصور

تسجّل <span dir="ltr">`media = document.add_media(source)`</span> صورة. يمكن أن يكون <span dir="ltr">`source`</span>:

- كائن IO، مثل <span dir="ltr">`File`</span> أو <span dir="ltr">`StringIO`</span>
- مسار ملف
- عنوان URL من نوع <span dir="ltr">`http`</span> أو <span dir="ltr">`https`</span> أو <span dir="ltr">`ftp`</span>

تُقرأ الصورة عند استدعاء <span dir="ltr">`render`</span>، وتُخزَّن دائمًا كجزء PNG، لذا فملفات PNG هي الخيار الآمن.

**الصور المضمّنة**: تضع <span dir="ltr">`p.image(media, options)`</span> الصورة في سطر النص.

| الخيار | القيم | التأثير |
|---|---|---|
| <span dir="ltr">`style`</span> | قاموس بخصائص CSS | الحجم، مثل <span dir="ltr">`{ width: "120pt", height: "60pt" }`</span> |
| <span dir="ltr">`title`</span> | سلسلة نصية | عنوان الصورة |
| <span dir="ltr">`fill_color`</span> | لون VML مثل <span dir="ltr">`"#00FF00"`</span> | تعبئة خلف الصورة |
| <span dir="ltr">`opacity`</span> | سلسلة نصية مثل <span dir="ltr">`"50%"`</span> | عتامة التعبئة |
| <span dir="ltr">`wrap`</span> | <span dir="ltr">`true`</span> | يوقف التفاف النص حول الصورة |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**الصور المحددة الموضع**: تثبّت <span dir="ltr">`p.image(media, pH:, pV:, wrap: true, width:, height:)`</span> الصورة على إزاحة من العمود (<span dir="ltr">`pH`</span>) ومن الفقرة (<span dir="ltr">`pV`</span>)، بوحدة EMU. يجب تمرير <span dir="ltr">`pH`</span> و<span dir="ltr">`pV`</span> و<span dir="ltr">`wrap`</span> جميعًا، وإلا وُضعت الصورة مضمّنة في السطر.

**صور رأس الصفحة**: <span dir="ltr">`h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`</span>. القيمة <span dir="ltr">`wrap: "none"`</span> توقف الالتفاف، وأي قيمة أخرى تجعل النص يلتف حول الصورة بشكل مربع.

**مشكلة معروفة في الإصدار 0.1.0**: تحسب الصور المحددة الموضع وصور رؤوس الصفحات وصور خلايا الجداول عرضها وارتفاعها كليهما على أنه <span dir="ltr">`width * height * 14.92`</span> EMU، لذلك تظهر دائمًا مربعة وصغيرة جدًا في الغالب. أما الصور المضمّنة التي يُضبط حجمها عبر <span dir="ltr">`style:`</span> فلا تتأثر. راجع [أفضل الممارسات](Best-Practices#images) للاطلاع على حل بديل.

## <a id="headers-and-footers"></a>رؤوس الصفحات وتذييلاتها

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- يمكن أن يحتوي رأس الصفحة على <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`table`</span> و<span dir="ltr">`image`</span> و<span dir="ltr">`page_numbers`</span>، ويمكن أن يحتوي التذييل على <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`table`</span> و<span dir="ltr">`page_numbers`</span>.
- الخيارات المُمرَّرة إلى <span dir="ltr">`Header.new`</span> أو <span dir="ltr">`Footer.new`</span> تنطبق على نصوصها.
- يغطي <span dir="ltr">`add_header`</span> و<span dir="ltr">`add_footer`</span> الصفحات من 2 فصاعدًا، أما الصفحة 1 فتستخدم دائمًا <span dir="ltr">`add_first_header`</span> و<span dir="ltr">`add_first_footer`</span>. راجع [الإعدادات](Configuration#first-page).

## <a id="page-numbers"></a>أرقام الصفحات

تكتب <span dir="ltr">`page_numbers(name: "Page", align: "center")`</span> نص <span dir="ltr">`name`</span> يليه حقل PAGE. يظهر الرقم بخط Arial بحجم 10 نقاط وبلون رمادي داكن. القيمة الافتراضية لـ <span dir="ltr">`align`</span> هي <span dir="ltr">`"right"`</span>.

مشكلة معروفة: يُكتب رمز الحقل دون الشرطة المائلة العكسية في <span dir="ltr">`\* MERGEFORMAT`</span>.

## <a id="table-of-contents"></a>جدول المحتويات

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- التوقيع: <span dir="ltr">`add_row(text, page = nil, leader: "dot", pos: 9360)`</span>.
- عند تمرير رقم صفحة، يحصل السطر على علامة جدولة بمحاذاة إلى اليمين عند <span dir="ltr">`pos`</span> (بوحدة جزء من عشرين من النقطة، و9360 تساوي 6.5 بوصة) تُملأ بحرف <span dir="ltr">`leader`</span>. قيم حرف التعبئة مأخوذة من WordprocessingML، مثل <span dir="ltr">`"dot"`</span> و<span dir="ltr">`"hyphen"`</span> و<span dir="ltr">`"underscore"`</span> و<span dir="ltr">`"middleDot"`</span> و<span dir="ltr">`"none"`</span>.
- من دون رقم صفحة، يكون السطر هو النص بعد علامة جدولة.
- أرقام الصفحات هي الأرقام التي تمررها. لا تنشئ docxi حقل TOC في Word ولا تحسب الصفحات.

## <a id="hyperlinks"></a>الارتباطات التشعبية

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

تعيد <span dir="ltr">`add_hyperlink(url)`</span> معرّف علاقة، وتضع <span dir="ltr">`hyperlink(text, url, id)`</span> الارتباط. يأتي الهدف من العلاقة، لذا مرّر المعرّف المُعاد لعنوان URL الذي تريده. الارتباطات الخارجية فقط مدعومة.

## <a id="frames-text-boxes"></a>الإطارات (مربعات النص)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- الإطارات مربعات نص VML بتعبئة رمادية فاتحة شبه شفافة ثابتة وحدّ خارجي أسود. يلتف النص فوقها وتحتها.
- داخل الإطار يمكنك استدعاء <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`tab`</span> و<span dir="ltr">`image`</span> و<span dir="ltr">`frame`</span>.
- يضيف <span dir="ltr">`spacing: true`</span> مسافة قبل النص.

## <a id="output"></a>المخرجات

تعيد <span dir="ltr">`word.render`</span> كائن <span dir="ltr">`StringIO`</span> يحتوي على حزمة <span dir="ltr">`.docx`</span> كاملة. يُكتب الأرشيف دون سجلات Zip64، لذا يمكن لـ Google Docs وإصدارات LibreOffice القديمة فتحه. تسرد صفحة [البنية](Architecture) كل أجزاء الحزمة.

## <a id="related-pages"></a>صفحات ذات صلة

- [أمثلة الاستخدام](Usage-Examples)
- [مرجع API](API-Reference)
- [خارطة الطريق](Roadmap) للاطلاع على المشكلات المعروفة المخطط إصلاحها في الإصدار 0.2.0

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
