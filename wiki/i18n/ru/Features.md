<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Возможности docxi: абзацы, текст, таблицы, изображения и колонтитулы

На этой странице перечислены все элементы, которые docxi может добавить в документ Word, и параметры каждого из них. Примеры основаны на такой подготовке:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Ненужные параметры можно опускать. Цвета задаются шестнадцатеричными значениями без `#`, например `"8B0000"`. Единицы измерения и параметры страницы описаны в разделе [Настройка](Configuration).

## <a id="paragraphs"></a>Абзацы

`document.p(options) { |p| ... }` добавляет абзац.

| Параметр | Значения | Эффект |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Выравнивание. `"both"` выравнивает текст по ширине |
| `fill` | Шестнадцатеричный цвет | Заливка под абзацем и его текстом |
| `left`, `right` | Двадцатые доли пункта | Отступ. Задавайте оба, потому что один игнорируется |
| `ul` | `true` | Делает абзац элементом маркированного списка |
| `bottom` | `true` | Рисует границу под абзацем |

Параметры текста, такие как `size:`, `bold:` или `color:`, тоже можно передать в `p`. Они становятся значениями по умолчанию для каждого фрагмента текста в абзаце.

Внутри блока можно вызывать `text`, `br`, `tab`, `image`, `hyperlink` и `frame`. `document.text(string, options)` — короткий способ создать абзац с одним фрагментом.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Текст

`p.text(string, options)` добавляет фрагмент текста (run).

| Параметр | Значения | Эффект |
|---|---|---|
| `bold` | `true` | Полужирный |
| `italic` | `true` | Курсив |
| `underline` | `true` | Одинарное подчёркивание |
| `color` | Шестнадцатеричный цвет | Цвет текста |
| `size` | Пункты | Размер шрифта |
| `font` | Название шрифта | Шрифт, например `"Georgia"` |
| `fill` | Шестнадцатеричный цвет | Заливка под текстом |
| `space` | `true` | Сохраняет пробелы в начале и в конце текста |
| `br` | `true` | Добавляет разрыв строки после текста |

Строка `nil` пропускается.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Разрывы строк, разрывы страниц и табуляция

- `p.br` или `p.br(times: 3)` добавляет разрывы строк внутри абзаца.
- `document.br(page: true)` начинает новую страницу.
- `p.tab` или `p.tab(times: 2)` добавляет символы табуляции внутри абзаца.
- `document.tab(page: true)` добавляет абзац с единственной табуляцией.

Не вызывайте `document.br` или `document.tab` без `page: true`. Они помещают фрагмент текста прямо в тело документа, а это недопустимый WordprocessingML.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Таблицы

`document.table(options)` принимает блок, внутри которого `table.tr(options)` добавляет строки, а `tr.tc(options)` — ячейки.

**Параметры таблицы**

| Параметр | Значения | Эффект |
|---|---|---|
| `columns_width` | Массив чисел | Ширина столбцов, умноженная на 14.92 для перевода в двадцатые доли пункта. Также ширина ячеек по умолчанию |
| `height` | Двадцатые доли пункта | Минимальная высота каждой строки |
| `iwidth` | Двадцатые доли пункта | Отступ таблицы от левого поля |
| `style` | Идентификатор стиля | Стиль таблицы. docxi определяет `TableNormal` и `TableCustom` |
| `width` | Число | Предпочтительная ширина таблицы, записывается с типом `auto`. Без `columns_width` эту же ширину, умноженную на 14.92, получает и каждая ячейка |
| `align` | `"left"`, `"center"`, `"right"` | Выравнивание таблицы. Текст, добавленный через `tc.text`, выравнивается так же |
| `padding` | `true` | Добавляет фиксированные внутренние поля в ячейках |
| `borders` | Хеш | Известная проблема: применяется к каждой ячейке, а не к внешнему контуру таблицы |

**Параметры строки**: `height:` задаёт минимальную высоту строки. Любой параметр ячейки, переданный в `tr`, применяется ко всем ячейкам строки. Параметры таблицы так же передаются каждой строке и ячейке.

**Параметры ячейки**

| Параметр | Значения | Эффект |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Толщина границы в восьмых долях пункта. `nil` означает, что с этой стороны границы нет |
| `fill` | Шестнадцатеричный цвет | Заливка ячейки. Ячейки без заливки белые (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Вертикальное выравнивание |
| `colspan` | Целое число | Сколько столбцов занимает ячейка |
| `rowspan` | `true` | Начинает вертикальное объединение |
| `merged` | `true` | Продолжает вертикальное объединение ячейки сверху. Её содержимое игнорируется |
| `width` | Число | Ширина ячейки, умноженная на 14.92 |

**Содержимое ячейки**

- `tc.text(string, options)`: параметры ячейки и строки, например `bold:` или `align:`, применяются и к тексту.
- `tc.p(options) { |p| ... }`: абзац только со своими параметрами.
- `tc.image(media, width:, height:)`: изображение. О проблеме с размером см. раздел [Изображения](#images).
- `tc.page_numbers`: поле номера страницы. Известная проблема: перед ним выводится текст «GlobalOptions».

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

## <a id="images"></a>Изображения

`media = document.add_media(source)` регистрирует изображение. `source` может быть:

- объектом IO, например `File` или `StringIO`
- путём к файлу
- URL с протоколом `http`, `https` или `ftp`

Изображение читается при вызове `render` и всегда сохраняется как часть PNG, поэтому файлы PNG — надёжный выбор.

**Встроенные изображения**: `p.image(media, options)` размещает изображение в строке текста.

| Параметр | Значения | Эффект |
|---|---|---|
| `style` | Хеш свойств CSS | Размер, например `{ width: "120pt", height: "60pt" }` |
| `title` | Строка | Заголовок изображения |
| `fill_color` | Цвет VML, например `"#00FF00"` | Заливка под изображением |
| `opacity` | Строка, например `"50%"` | Непрозрачность заливки |
| `wrap` | `true` | Отключает обтекание изображения текстом |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Позиционированные изображения**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` привязывает изображение со смещением относительно столбца (`pH`) и абзаца (`pV`) в EMU. Нужно указать все три параметра `pH`, `pV` и `wrap`, иначе изображение размещается в строке.

**Изображения в верхнем колонтитуле**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` отключает обтекание. Любое другое значение включает обтекание текстом «вокруг рамки».

**Известная проблема в 0.1.0**: позиционированные изображения, изображения в колонтитулах и в ячейках таблиц вычисляют и ширину, и высоту как `width * height * 14.92` EMU. Они всегда получаются квадратными и обычно крошечными. Встроенных изображений, размер которых задан через `style:`, это не касается. Обходной путь описан в разделе [Рекомендации](Best-Practices#images).

## <a id="headers-and-footers"></a>Колонтитулы

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

- Верхний колонтитул может содержать `text`, `br`, `table`, `image` и `page_numbers`. Нижний колонтитул может содержать `text`, `br`, `table` и `page_numbers`.
- Параметры, переданные в `Header.new` или `Footer.new`, применяются к их тексту.
- `add_header` и `add_footer` действуют начиная со страницы 2. Для страницы 1 всегда используются `add_first_header` и `add_first_footer`. См. [Настройка](Configuration#first-page).

## <a id="page-numbers"></a>Номера страниц

`page_numbers(name: "Page", align: "center")` записывает текст из `name`, а за ним поле PAGE. Номер выводится шрифтом Arial 10 pt тёмно-серого цвета. По умолчанию `align` равен `"right"`.

Известная проблема: в коде поля не хватает обратной косой черты в `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Оглавление

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Если указан номер страницы, пункт получает позицию табуляции с выравниванием по правому краю в точке `pos` (в двадцатых долях пункта; 9360 — это 6,5 дюйма), заполненную символом из `leader`. Значения заполнителя берутся из WordprocessingML, например `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` и `"none"`.
- Без номера страницы пункт состоит из текста после табуляции.
- Номера страниц — это те, что вы передаёте. docxi не создаёт поле TOC Word и не вычисляет страницы.

## <a id="hyperlinks"></a>Гиперссылки

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` возвращает идентификатор связи, а `hyperlink(text, url, id)` вставляет ссылку. Адрес берётся из связи, поэтому передавайте идентификатор, возвращённый для нужного URL. Поддерживаются только внешние ссылки.

## <a id="frames-text-boxes"></a>Рамки (надписи)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Рамки — это надписи VML с фиксированной светло-серой полупрозрачной заливкой и чёрным контуром. Текст располагается выше и ниже них.
- Внутри рамки можно вызывать `text`, `br`, `tab`, `image` и `frame`.
- `spacing: true` добавляет отступ перед текстом.

## <a id="output"></a>Результат

`word.render` возвращает `StringIO` с полным пакетом `.docx`. Архив записывается без записей Zip64, поэтому его открывают Google Docs и старые версии LibreOffice. Все части пакета перечислены в разделе [Архитектура](Architecture).

## <a id="related-pages"></a>Связанные страницы

- [Примеры использования](Usage-Examples)
- [Справочник API](API-Reference)
- [Планы](Roadmap): известные проблемы, которые планируется исправить в 0.2.0

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
