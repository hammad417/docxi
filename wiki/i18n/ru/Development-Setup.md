<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Среда разработки: запуск тестов и инструментов docxi

Эта страница помогает запустить рабочую копию docxi и рассказывает о тестах, примерах документов и настройке CI.

## <a id="get-the-code"></a>Получите код

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Для разработки подходит любой Ruby начиная с 2.5. Зависимости для разработки — minitest и rake.

## <a id="rake-tasks"></a>Задачи Rake

| Команда | Что делает |
|---|---|
| `bundle exec rake test` | Запускает набор тестов (задача по умолчанию) |
| `bundle exec rake samples` | Записывает примеры документов в `tmp/samples/*.docx`, чтобы открыть их в Word, LibreOffice или Google Docs |
| `bundle exec rake golden` | Перезаписывает `test/golden` по текущему коду. Используйте только после намеренного изменения результата |
| `bundle exec rake build` | Собирает гем в `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Запуск тестов с загруженным Rails

В папке `gemfiles/` есть по одному Gemfile на каждую версию Rails, а также несколько файлов для старых версий rubyzip и nokogiri:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` загружает весь фреймворк перед тестами. Также он сравнивает правила `blank?` в docxi с правилами ActiveSupport.

## <a id="what-the-tests-cover"></a>Что проверяют тесты

| Файл | Проверяет |
|---|---|
| `test/test_golden_output.rb` | Каждая часть двух примеров документов совпадает с тем, что создавала 0.0.6 |
| `test/test_package.rb` | XML корректен, связи указывают на существующие части, у каждой части есть тип содержимого, и записей Zip64 нет |
| `test/test_media.rb` | Изображения из путей, объектов IO и URL, а также то, что строка с `\|` никогда не выполняется как команда |
| `test/test_fixed_features.rb` | `tab`, `frame` и номера страниц в верхних колонтитулах |
| `test/test_util.rb` | `Docxi::Util.blank?` следует правилам ActiveSupport |
| `test/test_rails_independence.rb` | docxi работает без ActiveSupport |
| `test/test_warnings.rb` | При формировании документа `lib/` не выдаёт предупреждений об устаревании |

Примеры документов находятся в `test/support/sample_documents.rb`, а `test/support/package.rb` читает `.docx` прямо из его zip-структуры.

## <a id="continuous-integration"></a>Непрерывная интеграция

`.github/workflows/ci.yml` запускается при каждом push в `master` и для каждого pull request. Он охватывает:

- Ruby 2.5, 2.6 и 4.0 без Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) и 8.1 (4.0)
- Rails `main`, которому разрешено падать
- rubyzip 1.3 и 2.4, а также nokogiri 1.11
- задание, которое конвертирует примеры документов в PDF с помощью LibreOffice

Выпуск версий описан в разделе [Развёртывание](Deployment).

## <a id="related-pages"></a>Связанные страницы

- [Участие в разработке](Contributing)
- [Архитектура](Architecture)
- [Развёртывание](Deployment)

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
