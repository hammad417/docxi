<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# docxi ile başlarken: Ruby ile ilk Word belgeniz

Bu rehber sizi boş bir klasörden biçimlendirilmiş bir `.docx` dosyasına kadar götürür. Temel Ruby bilgisine sahip olduğunuz varsayılır. Rails veya Microsoft Word gerekmez.

## <a id="1-install-the-gem"></a>1. Gem'i kurun

```sh
gem install docxi
```

Gemfile kullanan bir projedeyseniz bunun yerine `gem "docxi", "~> 0.1"` satırını ekleyip `bundle install` komutunu çalıştırın. Ayrıntılar [Kurulum](Installation) sayfasındadır.

## <a id="2-create-a-document"></a>2. Bir belge oluşturun

Şunu `hello.rb` adıyla kaydedin:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

`ruby hello.rb` komutunu (Bundler projesinde `bundle exec ruby hello.rb`) çalıştırın ve `hello.docx` dosyasını açın.

## <a id="3-what-each-line-does"></a>3. Her satır ne yapar

- `Docxi::Document.new`, varsayılan sayfa kenar boşluklarıyla boş bir belge oluşturur.
- `word.document`, belgenin gövdesidir. Eklediğiniz her şey bunun üzerinden geçer.
- `document.p` bir paragraf ekler. `p` için verilen seçenekler, örneğin `align:`, hem paragrafa hem de içindeki metne uygulanır.
- `p.text` bir metin parçası (run) ekler. `size:` punto cinsindendir.
- `word.render`, `.docx` paketini oluşturur ve bir `StringIO` döndürür. Bunun `.string` değeri dosyanın baytlarını içerir.

## <a id="4-add-a-list-and-a-table"></a>4. Bir liste ve bir tablo ekleyin

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

Dikkat edilecek iki nokta var:

- `space: true`, bir metin parçasının başındaki veya sonundaki boşluğu korur. Bu seçenek olmadan Word boşluğu siler ve kelimeler birbirine yapışır.
- `tr` üzerindeki seçenekler o satırdaki her hücreye uygulanır; bu yüzden başlık satırının tamamı gölgelendirilir ve kalın olur.

## <a id="5-where-to-go-next"></a>5. Sırada ne var

- [Kullanım örnekleri](Usage-Examples): raporlar, faturalar ve mektuplar için eksiksiz tarifler
- [Özellikler](Features): paragraflar, metin, tablolar, görseller, üstbilgi ve altbilgiler için tüm seçenekler
- [Entegrasyonlar](Integrations): belgeyi bir Rails controller'ından göndermek
- [En iyi uygulamalar](Best-Practices): yaygın sürprizleri baştan önlemek

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
