<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Справочник API docxi: классы и методы

docxi — библиотека для Ruby, поэтому на этой странице перечислен её Ruby API. HTTP API нет. Параметры каждого метода описаны в разделе [Возможности](Features), а классы — в разделе [Архитектура](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Метод | Возвращает | Примечания |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` — поля страницы. См. [Настройка](Configuration) |
| `#document` | `Docxi::Word::Document` | Тело, в которое вы добавляете содержимое |
| `#render` | `StringIO` | Полный пакет `.docx`. `.string` возвращает байты |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Метод | Возвращает | Добавляет |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Абзац |
| `text(string, options = {})` | Paragraph | Абзац с одним фрагментом текста |
| `br(page: true)` | Break | Разрыв страницы |
| `tab(page: true)` | Tab | Абзац с одной табуляцией |
| `table(options = {}) { \|table\| }` | Table | Таблицу |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Пункты оглавления |
| `frame(text, options = {})` | Paragraph | Абзац с надписью |
| `add_media(io_path_or_url, options = {})` | Media | Изображение, которое размещается через `image` |
| `add_hyperlink(url, options = {})` | String | Идентификатор связи для `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Верхний колонтитул со страницы 2 и для страницы 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Нижний колонтитул со страницы 2 и для страницы 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Абзац (`Docxi::Word::Contents::Paragraph`)

| Метод | Добавляет |
|---|---|
| `text(string, options = {})` | Фрагмент текста. Параметры абзаца служат значениями по умолчанию |
| `br(times: 1)` | Разрывы строк |
| `tab(times: 1)` | Символы табуляции |
| `image(media, options = {})` | Встроенное или позиционированное изображение |
| `hyperlink(text, url, id)` | Ссылку, использующую идентификатор связи `id` из `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Надпись |

## <a id="tables"></a>Таблицы

| Метод | Добавляет |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Строку. Параметры таблицы служат значениями по умолчанию |
| `TableRow#tc(options = {}) { \|tc\| }` | Ячейку. Параметры строки служат значениями по умолчанию |
| `TableCell#text(string, options = {})` | Абзац. Параметры ячейки служат значениями по умолчанию |
| `TableCell#p(options = {}) { \|p\| }` | Абзац со своими параметрами |
| `TableCell#image(media, width:, height:)` | Изображение |
| `TableCell#page_numbers` | Поле номера страницы |

## <a id="table-of-contents"></a>Оглавление

| Метод | Добавляет |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Один пункт оглавления |

## <a id="headers-and-footers"></a>Колонтитулы

| Класс | Методы |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Рамки

`Docxi::Word::Contents::Frame` принимает в своём блоке `text`, `br`, `tab`, `image` и `frame`.

## <a id="constants-and-helpers"></a>Константы и вспомогательные методы

| Имя | Значение |
|---|---|
| `Docxi::VERSION` | Установленная версия, например `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Записывает одну часть в архив, заранее задав её размер. Используется внутри для каждой части |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Внутренние вспомогательные методы с правилами `blank?` из ActiveSupport |

## <a id="related-pages"></a>Связанные страницы

- [Возможности](Features)
- [Примеры использования](Usage-Examples)
- [Архитектура](Architecture)

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
