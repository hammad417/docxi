<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API başvurusu: sınıflar ve yöntemler

docxi bir Ruby kütüphanesidir; bu yüzden bu sayfa Ruby API'sini listeler. HTTP API'si yoktur. Her yöntemin seçenekleri [Özellikler](Features) sayfasında, sınıflar ise [Mimari](Architecture) sayfasında açıklanır.

## <a id="docxidocument"></a>Docxi::Document

| Yöntem | Döndürdüğü | Notlar |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` sayfa kenar boşluklarıdır. [Yapılandırma](Configuration) sayfasına bakın |
| `#document` | `Docxi::Word::Document` | İçerik eklediğiniz gövde |
| `#render` | `StringIO` | Eksiksiz `.docx` paketi. `.string` baytları verir |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Yöntem | Döndürdüğü | Eklediği |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Bir paragraf |
| `text(string, options = {})` | Paragraph | Tek metin parçası içeren bir paragraf |
| `br(page: true)` | Break | Bir sayfa sonu |
| `tab(page: true)` | Tab | Tek sekme içeren bir paragraf |
| `table(options = {}) { \|table\| }` | Table | Bir tablo |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | İçindekiler girdileri |
| `frame(text, options = {})` | Paragraph | Metin kutusu içeren bir paragraf |
| `add_media(io_path_or_url, options = {})` | Media | `image` ile yerleştirilecek bir görsel |
| `add_hyperlink(url, options = {})` | String | `hyperlink` için bir ilişki kimliği |
| `add_header(header)`, `add_first_header(header)` | Header | 2. sayfadan itibaren ve 1. sayfa için üstbilgi |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | 2. sayfadan itibaren ve 1. sayfa için altbilgi |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Paragraf (`Docxi::Word::Contents::Paragraph`)

| Yöntem | Eklediği |
|---|---|
| `text(string, options = {})` | Bir metin parçası. Paragrafın seçenekleri bunun varsayılanlarıdır |
| `br(times: 1)` | Satır sonları |
| `tab(times: 1)` | Sekme karakterleri |
| `image(media, options = {})` | Satır içi veya konumlandırılmış bir görsel |
| `hyperlink(text, url, id)` | `add_hyperlink` tarafından döndürülen ilişki `id` değerini kullanan bir bağlantı |
| `frame(options = {}) { \|frame\| }` | Bir metin kutusu |

## <a id="tables"></a>Tablolar

| Yöntem | Eklediği |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Bir satır. Tablonun seçenekleri bunun varsayılanlarıdır |
| `TableRow#tc(options = {}) { \|tc\| }` | Bir hücre. Satırın seçenekleri bunun varsayılanlarıdır |
| `TableCell#text(string, options = {})` | Bir paragraf. Hücrenin seçenekleri bunun varsayılanlarıdır |
| `TableCell#p(options = {}) { \|p\| }` | Kendi seçenekleri olan bir paragraf |
| `TableCell#image(media, width:, height:)` | Bir görsel |
| `TableCell#page_numbers` | Bir sayfa numarası alanı |

## <a id="table-of-contents"></a>İçindekiler tablosu

| Yöntem | Eklediği |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Bir içindekiler girdisi |

## <a id="headers-and-footers"></a>Üstbilgi ve altbilgiler

| Sınıf | Yöntemler |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Çerçeveler

`Docxi::Word::Contents::Frame`, bloğunun içinde `text`, `br`, `tab`, `image` ve `frame` kabul eder.

## <a id="constants-and-helpers"></a>Sabitler ve yardımcılar

| Ad | Anlamı |
|---|---|
| `Docxi::VERSION` | Kurulu sürüm, örneğin `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Boyutunu önceden ayarlayarak arşive tek bir parça yazar. Her parça tarafından dahili olarak kullanılır |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | ActiveSupport'un `blank?` kurallarını izleyen dahili yardımcılar |

## <a id="related-pages"></a>İlgili sayfalar

- [Özellikler](Features)
- [Kullanım örnekleri](Usage-Examples)
- [Mimari](Architecture)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
