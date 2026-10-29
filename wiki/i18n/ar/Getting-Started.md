<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# البدء مع docxi: أول مستند Word لك باستخدام Ruby

يأخذك هذا الدليل من مجلد فارغ إلى ملف <span dir="ltr">`.docx`</span> منسّق. يفترض الدليل أنك تعرف أساسيات Ruby، ولا تحتاج إلى Rails أو Microsoft Word.

## <a id="1-install-the-gem"></a>1. ثبّت الحزمة

```sh
gem install docxi
```

في مشروع يحتوي على Gemfile، أضف <span dir="ltr">`gem "docxi", "~> 0.1"`</span> وشغّل <span dir="ltr">`bundle install`</span> بدلًا من ذلك. تجد التفاصيل في صفحة [التثبيت](Installation).

## <a id="2-create-a-document"></a>2. أنشئ مستندًا

احفظ ما يلي في ملف باسم <span dir="ltr">`hello.rb`</span>:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

شغّل <span dir="ltr">`ruby hello.rb`</span> (أو <span dir="ltr">`bundle exec ruby hello.rb`</span> في مشروع يستخدم Bundler) ثم افتح <span dir="ltr">`hello.docx`</span>.

## <a id="3-what-each-line-does"></a>3. ما الذي يفعله كل سطر

- ينشئ <span dir="ltr">`Docxi::Document.new`</span> مستندًا فارغًا بهوامش الصفحة الافتراضية.
- يمثّل <span dir="ltr">`word.document`</span> متن المستند، وكل ما تضيفه يمر عبره.
- يضيف <span dir="ltr">`document.p`</span> فقرة. الخيارات التي تمررها إلى <span dir="ltr">`p`</span>، مثل <span dir="ltr">`align:`</span>، تنطبق على الفقرة وعلى النص الذي بداخلها.
- يضيف <span dir="ltr">`p.text`</span> مقطعًا نصيًا (run). وحدة <span dir="ltr">`size:`</span> هي النقطة.
- يبني <span dir="ltr">`word.render`</span> حزمة <span dir="ltr">`.docx`</span> ويعيد كائن <span dir="ltr">`StringIO`</span>، وتحتوي قيمته <span dir="ltr">`.string`</span> على بايتات الملف.

## <a id="4-add-a-list-and-a-table"></a>4. أضف قائمة وجدولًا

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

أمران يجب الانتباه إليهما:

- يحافظ الخيار <span dir="ltr">`space: true`</span> على المسافة في بداية المقطع أو نهايته. من دونه يحذف Word المسافة فتلتصق الكلمات ببعضها.
- الخيارات المحددة على <span dir="ltr">`tr`</span> تنطبق على كل خلية في ذلك الصف، ولذلك يظهر صف العناوين كله مظللًا وبخط غامق.

## <a id="5-where-to-go-next"></a>5. الخطوات التالية

- [أمثلة الاستخدام](Usage-Examples): وصفات كاملة للتقارير والفواتير والخطابات
- [الميزات](Features): كل خيارات الفقرات والنصوص والجداول والصور ورؤوس الصفحات وتذييلاتها
- [التكامل](Integrations): إرسال المستند من وحدة تحكم في Rails
- [أفضل الممارسات](Best-Practices): تجنّب المفاجآت الشائعة قبل حدوثها

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
