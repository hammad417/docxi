<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# docxi yapılandırması: sayfa kenar boşlukları, sayfa yapısı ve birimler

docxi'nin yapılandırma dosyası, ortam değişkenleri veya genel ayarları yoktur. Her şey belge başına ya da öğe başına ayarlanır. Bu sayfa belge düzeyindeki ayarları ve seçeneklerin kullandığı birimleri anlatır. Her öğenin seçenekleri [Özellikler](Features) sayfasında listelenir.

## <a id="page-margins"></a>Sayfa kenar boşlukları

Kenar boşlukları belge oluşturulurken verilir:

```ruby
require "docxi"

# Tüm değerler puntonun yirmide biri cinsindendir (1440 = 1 inç).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Seçenek | Varsayılan | Anlamı |
|---|---|---|
| `top` | `720` (0,5 inç) | Üst kenar boşluğu |
| `right` | `1440` (1 inç) | Sağ kenar boşluğu |
| `bottom` | `720` (0,5 inç) | Alt kenar boşluğu |
| `left` | `1440` (1 inç) | Sol kenar boşluğu |
| `header` | `190` | Sayfanın üst kenarından üstbilgiye olan uzaklık |
| `footer` | `200` | Sayfanın alt kenarından altbilgiye olan uzaklık |

Verdiğiniz hash varsayılanların yerine geçer; bu yüzden `top`, `right`, `bottom` ve `left` değerlerini birlikte ayarlayın. Bu dördünden atladığınız her biri boş bir değer olarak yazılır. `header` ve `footer` eksik veya boş olduğunda 190 ve 200 kullanılır.

## <a id="page-size-and-orientation"></a>Sayfa boyutu ve yönü

Sayfalar dikey yönde US Letter boyutundadır: puntonun yirmide biri cinsinden 12240 × 15840, yani 8,5 × 11 inç. 0.1.0'da boyut ve yön değiştirilemez.

## <a id="first-page"></a>İlk sayfa

docxi, Word'ün "İlk sayfa farklı" ayarını her zaman açar. `add_header` ve `add_footer` ile eklenen üstbilgi ve altbilgiler 2. sayfadan itibaren kullanılır. 1. sayfa `add_first_header` ve `add_first_footer` kullanır; ilk sayfada üstbilgi veya altbilgi gerekiyorsa bunları da ayarlayın. [Özellikler](Features#headers-and-footers) sayfasına bakın.

## <a id="what-every-document-contains"></a>Her belgede neler bulunur

- Varsayılan yazı tipi Arial, 12 pt'dir.
- Tanımlı stiller: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` ve `NoList`.
- Madde işaretli öğeler (`ul: true`) madde işareti olarak tire kullanır.
- Belge özellikleri, uygulama adını "Docxi Ruby Gem" olarak ve belgenin oluşturulduğu zamanı kaydeder.

## <a id="units"></a>Birimler

| Seçenek | Birim |
|---|---|
| Metnin `size:` değeri | Punto (docxi bunu iki katına çıkararak Word'ün yarım punto birimine çevirir) |
| Kenar boşlukları, paragrafın `left:` ve `right:` değerleri, tablonun `iwidth:` değeri, satırın `height:` değeri | Puntonun yirmide biri (inç başına 1440) |
| Hücre `borders:` kalınlıkları | Puntonun sekizde biri |
| `columns_width:` ve hücrenin `width:` değeri | Puntonun yirmide birini elde etmek için 14.92 ile çarpılır; 96 dpi'de piksele yakındır |
| Satır içi görselin `style:` değeri | `"120pt"` gibi CSS uzunlukları |
| Konumlandırılmış görselin `pH:` ve `pV:` değerleri | EMU (inç başına 914400) |
| Metin ve dolgu renkleri | `#` olmadan onaltılık değer, örneğin `"8B0000"` |
| Görselin `fill_color:` değeri | `#` içeren VML rengi, örneğin `"#00FF00"` |

## <a id="related-pages"></a>İlgili sayfalar

- [Özellikler](Features)
- [API başvurusu](API-Reference)
- [En iyi uygulamalar](Best-Practices)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
