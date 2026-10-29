<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi özellikleri: paragraflar, metin, tablolar, görseller, üstbilgi ve altbilgiler

Bu sayfa, docxi'nin bir Word belgesine ekleyebildiği her öğeyi ve her birinin kabul ettiği seçenekleri listeler. Örnekler şu hazırlık üzerine kuruludur:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

İhtiyacınız olmayan seçenekleri atlayabilirsiniz. Renkler `#` olmadan onaltılık değerlerdir, örneğin `"8B0000"`. Birimler ve sayfa yapısı için [Yapılandırma](Configuration) sayfasına bakın.

## <a id="paragraphs"></a>Paragraflar

`document.p(options) { |p| ... }` bir paragraf ekler.

| Seçenek | Değerler | Etkisi |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Hizalama. `"both"` metni iki yana yaslar |
| `fill` | Onaltılık renk | Paragrafın ve metninin arkasındaki gölgelendirme |
| `left`, `right` | Puntonun yirmide biri | Girinti. İkisini birden ayarlayın, çünkü tek başına biri yok sayılır |
| `ul` | `true` | Paragrafı madde işaretli liste öğesi yapar |
| `bottom` | `true` | Paragrafın altına bir kenarlık çizer |

`size:`, `bold:` veya `color:` gibi metin seçenekleri de `p` için verilebilir. Bunlar paragraftaki her metin parçasının varsayılanı olur.

Bloğun içinde `text`, `br`, `tab`, `image`, `hyperlink` ve `frame` çağrılabilir. `document.text(string, options)`, tek bir metin parçası içeren bir paragrafın kısayoludur.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Metin

`p.text(string, options)` bir metin parçası (run) ekler.

| Seçenek | Değerler | Etkisi |
|---|---|---|
| `bold` | `true` | Kalın |
| `italic` | `true` | İtalik |
| `underline` | `true` | Tek çizgi altı çizili |
| `color` | Onaltılık renk | Metin rengi |
| `size` | Punto | Yazı tipi boyutu |
| `font` | Yazı tipi adı | Yazı tipi, örneğin `"Georgia"` |
| `fill` | Onaltılık renk | Metnin arkasındaki gölgelendirme |
| `space` | `true` | Metnin başındaki ve sonundaki boşlukları korur |
| `br` | `true` | Metnin ardından satır sonu ekler |

`nil` dize atlanır.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Satır sonları, sayfa sonları ve sekmeler

- `p.br` veya `p.br(times: 3)`, paragrafın içine satır sonları ekler.
- `document.br(page: true)` yeni bir sayfa başlatır.
- `p.tab` veya `p.tab(times: 2)`, paragrafın içine sekme karakterleri ekler.
- `document.tab(page: true)`, tek bir sekme içeren bir paragraf ekler.

`document.br` veya `document.tab` yöntemini `page: true` olmadan çağırmayın. Bu yöntemler metin parçasını doğrudan belge gövdesine koyar ve bu geçerli bir WordprocessingML değildir.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tablolar

`document.table(options)` bir blok alır; bloğun içinde `table.tr(options)` satır, `tr.tc(options)` ise hücre ekler.

**Tablo seçenekleri**

| Seçenek | Değerler | Etkisi |
|---|---|---|
| `columns_width` | Sayı dizisi | Sütun genişlikleri; 14.92 ile çarpılarak puntonun yirmide birine çevrilir. Aynı zamanda hücrelerin varsayılan genişlikleridir |
| `height` | Puntonun yirmide biri | Her satırın en düşük yüksekliği |
| `iwidth` | Puntonun yirmide biri | Tablonun sol kenar boşluğundan girintisi |
| `style` | Stil kimliği | Tablo stili. docxi, `TableNormal` ve `TableCustom` stillerini tanımlar |
| `width` | Sayı | Tercih edilen tablo genişliği, `auto` türüyle yazılır. `columns_width` yoksa her hücre de bu genişliği 14.92 ile çarpılmış olarak alır |
| `align` | `"left"`, `"center"`, `"right"` | Tablo hizalaması. `tc.text` ile eklenen metin de aynı şekilde hizalanır |
| `padding` | `true` | Hücrelerin içine sabit iç boşluk ekler |
| `borders` | Hash | Bilinen sorun: tablonun dış kenarına değil, her hücreye uygulanır |

**Satır seçenekleri**: `height:` satırın en düşük yüksekliğini ayarlar. `tr` için verilen her hücre seçeneği o satırdaki tüm hücrelere uygulanır. Tablo seçenekleri de aynı şekilde her satıra ve hücreye aktarılır.

**Hücre seçenekleri**

| Seçenek | Değerler | Etkisi |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Puntonun sekizde biri cinsinden kenarlık kalınlığı. `nil` o kenarda kenarlık olmadığı anlamına gelir |
| `fill` | Onaltılık renk | Hücre gölgelendirmesi. Dolgusu olmayan hücreler beyazdır (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Dikey hizalama |
| `colspan` | Tam sayı | Hücrenin kapladığı sütun sayısı |
| `rowspan` | `true` | Dikey birleştirmeyi başlatır |
| `merged` | `true` | Üstteki hücrenin dikey birleştirmesini sürdürür. İçeriği yok sayılır |
| `width` | Sayı | Hücre genişliği, 14.92 ile çarpılır |

**Hücre içeriği**

- `tc.text(string, options)`: `bold:` veya `align:` gibi hücre ve satır seçenekleri metne de uygulanır.
- `tc.p(options) { |p| ... }`: yalnızca kendi seçeneklerini kullanan bir paragraf.
- `tc.image(media, width:, height:)`: bir görsel. Boyut sorunu için [Görseller](#images) bölümüne bakın.
- `tc.page_numbers`: bir sayfa numarası alanı. Bilinen sorun: önüne "GlobalOptions" metni eklenir.

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

## <a id="images"></a>Görseller

`media = document.add_media(source)` bir görsel kaydeder. `source` şunlardan biri olabilir:

- `File` veya `StringIO` gibi bir IO nesnesi
- bir dosya yolu
- bir `http`, `https` veya `ftp` URL'si

Görsel, `render` çağrıldığında okunur ve her zaman PNG parçası olarak saklanır; bu yüzden en güvenli seçim PNG dosyalarıdır.

**Satır içi görseller**: `p.image(media, options)` görseli metin satırının içine yerleştirir.

| Seçenek | Değerler | Etkisi |
|---|---|---|
| `style` | CSS özelliklerinden oluşan hash | Boyut, örneğin `{ width: "120pt", height: "60pt" }` |
| `title` | Dize | Görselin başlığı |
| `fill_color` | `"#00FF00"` gibi VML rengi | Görselin arkasındaki dolgu |
| `opacity` | `"50%"` gibi bir dize | Dolgunun opaklığı |
| `wrap` | `true` | Görselin çevresindeki metin kaydırmayı kapatır |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Konumlandırılmış görseller**: `p.image(media, pH:, pV:, wrap: true, width:, height:)`, görseli sütundan (`pH`) ve paragraftan (`pV`) EMU cinsinden bir uzaklıkla sabitler. `pH`, `pV` ve `wrap` değerlerinin hepsi verilmelidir; aksi halde görsel satır içine yerleştirilir.

**Üstbilgi görselleri**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` metin kaydırmayı kapatır. Başka herhangi bir değer metni görselin çevresinde kare biçiminde kaydırır.

**0.1.0'da bilinen sorun**: konumlandırılmış görseller, üstbilgi görselleri ve tablo hücresi görselleri genişliklerini ve yüksekliklerini `width * height * 14.92` EMU olarak hesaplar. Her zaman kare ve genellikle çok küçük çıkarlar. Boyutu `style:` ile verilen satır içi görseller bundan etkilenmez. Geçici çözüm için [En iyi uygulamalar](Best-Practices#images) sayfasına bakın.

## <a id="headers-and-footers"></a>Üstbilgi ve altbilgiler

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

- Üstbilgi `text`, `br`, `table`, `image` ve `page_numbers` içerebilir. Altbilgi `text`, `br`, `table` ve `page_numbers` içerebilir.
- `Header.new` veya `Footer.new` için verilen seçenekler bunların metnine uygulanır.
- `add_header` ve `add_footer`, 2. sayfa ve sonrasını kapsar. 1. sayfa her zaman `add_first_header` ve `add_first_footer` kullanır. [Yapılandırma](Configuration#first-page) sayfasına bakın.

## <a id="page-numbers"></a>Sayfa numaraları

`page_numbers(name: "Page", align: "center")`, `name` metnini ve ardından bir PAGE alanını yazar. Numara koyu gri renkte, Arial 10 pt ile gösterilir. `align` için varsayılan değer `"right"` olur.

Bilinen sorun: alan kodunda `\* MERGEFORMAT` içindeki ters eğik çizgi eksik yazılır.

## <a id="table-of-contents"></a>İçindekiler tablosu

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Sayfa numarası verildiğinde girdi, `pos` konumunda (puntonun yirmide biri cinsinden; 9360, 6,5 inç demektir) sağa hizalı bir sekme durağı alır ve bu alan `leader` karakteriyle doldurulur. Öncü karakter değerleri WordprocessingML'den gelir; örneğin `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` ve `"none"`.
- Sayfa numarası verilmediğinde girdi, bir sekmeden sonra gelen metinden oluşur.
- Sayfa numaraları sizin verdiklerinizdir. docxi bir Word TOC alanı oluşturmaz ve sayfaları hesaplamaz.

## <a id="hyperlinks"></a>Köprüler

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` bir ilişki kimliği döndürür, `hyperlink(text, url, id)` ise bağlantıyı yerleştirir. Hedef ilişkiden alınır; bu yüzden istediğiniz URL için döndürülen kimliği verin. Yalnızca dış bağlantılar desteklenir.

## <a id="frames-text-boxes"></a>Çerçeveler (metin kutuları)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Çerçeveler, sabit açık gri ve yarı saydam bir dolguya ve siyah bir dış çizgiye sahip VML metin kutularıdır. Metin bunların üstüne ve altına yerleşir.
- Bir çerçevenin içinde `text`, `br`, `tab`, `image` ve `frame` çağrılabilir.
- `spacing: true`, metinden önce boşluk ekler.

## <a id="output"></a>Çıktı

`word.render`, eksiksiz `.docx` paketini içeren bir `StringIO` döndürür. Arşiv Zip64 kayıtları olmadan yazılır; böylece Google Docs ve LibreOffice'in eski sürümleri onu açabilir. [Mimari](Architecture) sayfası paketin her parçasını listeler.

## <a id="related-pages"></a>İlgili sayfalar

- [Kullanım örnekleri](Usage-Examples)
- [API başvurusu](API-Reference)
- [Yol haritası](Roadmap): 0.2.0 için planlanan bilinen sorunlar

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
