<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# مرجع API لمكتبة docxi: الأصناف والدوال

مكتبة docxi مكتبة Ruby، لذا تسرد هذه الصفحة واجهتها البرمجية في Ruby، ولا توجد واجهة HTTP. خيارات كل دالة موصوفة في صفحة [الميزات](Features)، والأصناف مشروحة في صفحة [البنية](Architecture).

## <a id="docxidocument"></a>الصنف Docxi::Document

| الدالة | القيمة المُعادة | ملاحظات |
|---|---|---|
| <span dir="ltr">`Docxi::Document.new(options = {})`</span> | كائن Document | الخيارات <span dir="ltr">`options`</span> هي هوامش الصفحة. راجع [الإعدادات](Configuration) |
| <span dir="ltr">`#document`</span> | <span dir="ltr">`Docxi::Word::Document`</span> | المتن الذي تضيف إليه المحتوى |
| <span dir="ltr">`#render`</span> | <span dir="ltr">`StringIO`</span> | حزمة <span dir="ltr">`.docx`</span> الكاملة. تعيد <span dir="ltr">`.string`</span> البايتات |

## <a id="docxiworddocument-worddocument"></a>الصنف Docxi::Word::Document ‏(<span dir="ltr">`word.document`</span>)

| الدالة | القيمة المُعادة | ما تضيفه |
|---|---|---|
| <span dir="ltr">`p(options = {}) { \|p\| }`</span> | كائن Paragraph | فقرة |
| <span dir="ltr">`text(string, options = {})`</span> | كائن Paragraph | فقرة بمقطع نصي واحد |
| <span dir="ltr">`br(page: true)`</span> | كائن Break | فاصل صفحات |
| <span dir="ltr">`tab(page: true)`</span> | كائن Tab | فقرة تحتوي على علامة جدولة واحدة |
| <span dir="ltr">`table(options = {}) { \|table\| }`</span> | كائن Table | جدول |
| <span dir="ltr">`table_of_content(options = {}) { \|toc\| }`</span> | كائن TableOfContent | أسطر جدول المحتويات |
| <span dir="ltr">`frame(text, options = {})`</span> | كائن Paragraph | فقرة تحتوي على مربع نص |
| <span dir="ltr">`add_media(io_path_or_url, options = {})`</span> | كائن Media | صورة تضعها باستخدام <span dir="ltr">`image`</span> |
| <span dir="ltr">`add_hyperlink(url, options = {})`</span> | سلسلة نصية (String) | معرّف علاقة لاستخدامه مع <span dir="ltr">`hyperlink`</span> |
| <span dir="ltr">`add_header(header)`</span> و<span dir="ltr">`add_first_header(header)`</span> | كائن Header | رأس الصفحات من 2 فصاعدًا، ورأس الصفحة 1 |
| <span dir="ltr">`add_footer(footer)`</span> و<span dir="ltr">`add_first_footer(footer)`</span> | كائن Footer | تذييل الصفحات من 2 فصاعدًا، وتذييل الصفحة 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>الفقرة (<span dir="ltr">`Docxi::Word::Contents::Paragraph`</span>)

| الدالة | ما تضيفه |
|---|---|
| <span dir="ltr">`text(string, options = {})`</span> | مقطع نصي. خيارات الفقرة هي قيمه الافتراضية |
| <span dir="ltr">`br(times: 1)`</span> | فواصل أسطر |
| <span dir="ltr">`tab(times: 1)`</span> | علامات جدولة |
| <span dir="ltr">`image(media, options = {})`</span> | صورة مضمّنة أو محددة الموضع |
| <span dir="ltr">`hyperlink(text, url, id)`</span> | ارتباط يستخدم معرّف العلاقة <span dir="ltr">`id`</span> الذي أعادته <span dir="ltr">`add_hyperlink`</span> |
| <span dir="ltr">`frame(options = {}) { \|frame\| }`</span> | مربع نص |

## <a id="tables"></a>الجداول

| الدالة | ما تضيفه |
|---|---|
| <span dir="ltr">`Table#tr(options = {}) { \|tr\| }`</span> | صف. خيارات الجدول هي قيمه الافتراضية |
| <span dir="ltr">`TableRow#tc(options = {}) { \|tc\| }`</span> | خلية. خيارات الصف هي قيمها الافتراضية |
| <span dir="ltr">`TableCell#text(string, options = {})`</span> | فقرة. خيارات الخلية هي قيمها الافتراضية |
| <span dir="ltr">`TableCell#p(options = {}) { \|p\| }`</span> | فقرة بخياراتها الخاصة |
| <span dir="ltr">`TableCell#image(media, width:, height:)`</span> | صورة |
| <span dir="ltr">`TableCell#page_numbers`</span> | حقل رقم الصفحة |

## <a id="table-of-contents"></a>جدول المحتويات

| الدالة | ما تضيفه |
|---|---|
| <span dir="ltr">`TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)`</span> | سطر واحد في جدول المحتويات |

## <a id="headers-and-footers"></a>رؤوس الصفحات وتذييلاتها

| الصنف | الدوال |
|---|---|
| <span dir="ltr">`Docxi::Word::Headers::Header.new(options = {}) { \|h\| }`</span> | الدوال <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`table`</span> و<span dir="ltr">`image`</span> و<span dir="ltr">`page_numbers(options = {})`</span> |
| <span dir="ltr">`Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }`</span> | الدوال <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`table`</span> و<span dir="ltr">`page_numbers(options = {})`</span> |

## <a id="frames"></a>الإطارات

يقبل <span dir="ltr">`Docxi::Word::Contents::Frame`</span> داخل كتلته الدوال <span dir="ltr">`text`</span> و<span dir="ltr">`br`</span> و<span dir="ltr">`tab`</span> و<span dir="ltr">`image`</span> و<span dir="ltr">`frame`</span>.

## <a id="constants-and-helpers"></a>الثوابت والدوال المساعدة

| الاسم | المعنى |
|---|---|
| <span dir="ltr">`Docxi::VERSION`</span> | الإصدار المثبّت، مثل <span dir="ltr">`"0.1.0"`</span> |
| <span dir="ltr">`Docxi.write_entry(zip, name, data)`</span> | يكتب جزءًا واحدًا في الأرشيف بعد ضبط حجمه مسبقًا. تستخدمه كل الأجزاء داخليًا |
| <span dir="ltr">`Docxi::Util.blank?`</span> و<span dir="ltr">`Docxi::Util.present?`</span> | دوال مساعدة داخلية تتبع قواعد <span dir="ltr">`blank?`</span> في ActiveSupport |

## <a id="related-pages"></a>صفحات ذات صلة

- [الميزات](Features)
- [أمثلة الاستخدام](Usage-Examples)
- [البنية](Architecture)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
