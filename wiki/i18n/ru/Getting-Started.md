<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Начало работы с docxi: первый документ Word на Ruby

Это руководство проведёт вас от пустой папки до отформатированного файла `.docx`. Предполагается, что вы знаете основы Ruby. Rails и Microsoft Word не нужны.

## <a id="1-install-the-gem"></a>1. Установите гем

```sh
gem install docxi
```

Если в проекте есть Gemfile, вместо этого добавьте `gem "docxi", "~> 0.1"` и выполните `bundle install`. Подробности — на странице [Установка](Installation).

## <a id="2-create-a-document"></a>2. Создайте документ

Сохраните этот код в файл `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Выполните `ruby hello.rb` (или `bundle exec ruby hello.rb` в проекте с Bundler) и откройте `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Что делает каждая строка

- `Docxi::Document.new` создаёт пустой документ с полями страницы по умолчанию.
- `word.document` — это тело документа. Всё, что вы добавляете, проходит через него.
- `document.p` добавляет абзац. Параметры, переданные в `p`, например `align:`, применяются к абзацу и к тексту внутри него.
- `p.text` добавляет фрагмент текста (run). `size:` задаётся в пунктах.
- `word.render` собирает пакет `.docx` и возвращает `StringIO`. Его `.string` содержит байты файла.

## <a id="4-add-a-list-and-a-table"></a>4. Добавьте список и таблицу

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

Обратите внимание на две вещи:

- `space: true` сохраняет пробел в начале или в конце фрагмента. Без этого параметра Word удаляет пробел, и слова сливаются.
- Параметры `tr` применяются к каждой ячейке строки, поэтому вся строка заголовка получает заливку и полужирное начертание.

## <a id="5-where-to-go-next"></a>5. Что дальше

- [Примеры использования](Usage-Examples): полные рецепты для отчётов, счетов и писем
- [Возможности](Features): все параметры абзацев, текста, таблиц, изображений и колонтитулов
- [Интеграции](Integrations): отправка документа из контроллера Rails
- [Рекомендации](Best-Practices): как заранее избежать типичных сюрпризов

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
