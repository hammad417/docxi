<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->

<div dir="rtl">

# إعداد بيئة التطوير: تشغيل اختبارات docxi وأدواتها

تساعدك هذه الصفحة على تشغيل نسخة عمل من docxi، وتشرح الاختبارات والمستندات النموذجية وإعداد CI.

## <a id="get-the-code"></a>الحصول على الشيفرة

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

يصلح أي إصدار من Ruby بدءًا من 2.5 للتطوير. اعتماديات التطوير هي minitest وrake.

## <a id="rake-tasks"></a>مهام Rake

| الأمر | ما يفعله |
|---|---|
| <span dir="ltr">`bundle exec rake test`</span> | يشغّل مجموعة الاختبارات (المهمة الافتراضية) |
| <span dir="ltr">`bundle exec rake samples`</span> | يكتب المستندات النموذجية في <span dir="ltr">`tmp/samples/*.docx`</span> لفتحها في Word أو LibreOffice أو Google Docs |
| <span dir="ltr">`bundle exec rake golden`</span> | يعيد كتابة <span dir="ltr">`test/golden`</span> من الشيفرة الحالية. استخدمه فقط بعد تغيير مقصود في المخرجات |
| <span dir="ltr">`bundle exec rake build`</span> | يبني الحزمة في <span dir="ltr">`pkg/`</span> |

## <a id="running-the-tests-with-rails-loaded"></a>تشغيل الاختبارات مع تحميل Rails

يحتوي المجلد <span dir="ltr">`gemfiles/`</span> على ملف Gemfile لكل إصدار من Rails، إضافة إلى بعض الملفات لإصدارات أقدم من rubyzip وnokogiri:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

يحمّل <span dir="ltr">`DOCXI_WITH_RAILS=1`</span> إطار العمل كاملًا قبل الاختبارات، ويقارن أيضًا قواعد <span dir="ltr">`blank?`</span> في docxi بقواعد ActiveSupport.

## <a id="what-the-tests-cover"></a>ما تغطيه الاختبارات

| الملف | ما يتحقق منه |
|---|---|
| <span dir="ltr">`test/test_golden_output.rb`</span> | كل جزء من المستندين النموذجيين يطابق ما ولّده الإصدار 0.0.6 |
| <span dir="ltr">`test/test_package.rb`</span> | سلامة بنية XML، وإشارة العلاقات إلى أجزاء موجودة، ووجود نوع محتوى لكل جزء، وعدم وجود سجلات Zip64 |
| <span dir="ltr">`test/test_media.rb`</span> | الصور من المسارات وكائنات IO وعناوين URL، وأن السلسلة التي تبدأ بـ <span dir="ltr">`\|`</span> لا تُنفَّذ أبدًا كأمر |
| <span dir="ltr">`test/test_fixed_features.rb`</span> | الدوال <span dir="ltr">`tab`</span> و<span dir="ltr">`frame`</span> وأرقام الصفحات في رؤوس الصفحات |
| <span dir="ltr">`test/test_util.rb`</span> | اتباع <span dir="ltr">`Docxi::Util.blank?`</span> لقواعد ActiveSupport |
| <span dir="ltr">`test/test_rails_independence.rb`</span> | عمل docxi دون ActiveSupport |
| <span dir="ltr">`test/test_warnings.rb`</span> | عدم صدور أي تحذيرات إهمال (deprecation) من <span dir="ltr">`lib/`</span> أثناء التوليد |

توجد المستندات النموذجية في <span dir="ltr">`test/support/sample_documents.rb`</span>، ويقرأ <span dir="ltr">`test/support/package.rb`</span> ملف <span dir="ltr">`.docx`</span> مباشرة من بنى zip الخاصة به.

## <a id="continuous-integration"></a>التكامل المستمر

يعمل <span dir="ltr">`.github/workflows/ci.yml`</span> مع كل دفع إلى <span dir="ltr">`master`</span> ومع كل طلب سحب، ويغطي:

- إصدارات Ruby 2.5 و2.6 و4.0 دون Rails
- إصدارات Rails 5.2 ‏(Ruby 2.5) و6.0 ‏(2.7) و6.1 ‏(3.0) و7.0 ‏(3.1) و7.1 ‏(3.2) و7.2 ‏(3.3) و8.0 ‏(3.4) و8.1 ‏(4.0)
- الفرع <span dir="ltr">`main`</span> من Rails، ويُسمح له بالفشل
- مكتبة rubyzip 1.3 و2.4، ومكتبة nokogiri 1.11
- مهمة تحوّل المستندات النموذجية إلى PDF باستخدام LibreOffice

الإصدارات مشروحة في صفحة [النشر](Deployment).

## <a id="related-pages"></a>صفحات ذات صلة

- [المساهمة](Contributing)
- [البنية](Architecture)
- [النشر](Deployment)

---

[→ الصفحة الرئيسية للويكي](Home) · [README](https://github.com/hammad417/docxi#readme)

</div>
