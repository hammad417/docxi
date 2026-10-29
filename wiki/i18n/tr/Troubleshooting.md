<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi sorun giderme: kurulum hataları, çalışma zamanı hataları ve bozuk çıktı

Belirtinizi aşağıda bulun. Eski uygulamalardaki sorunların çoğu docxi 0.0.x kaynaklıdır ve 0.1.0'a yükseltmek bunları kod değişikliği gerekmeden çözer.

## <a id="installation"></a>Kurulum

**Bundler, Rails 6+ uygulamasında uyumlu `nokogiri` sürümleri bulamıyor.**
docxi 0.0.6, `nokogiri < 1.7` ile sabitlenmiştir; oysa 6.0'dan beri her Rails sürümü nokogiri 1.8.5 veya daha yenisini gerektirir. `gem "docxi", "~> 0.1"` kullanın.

**Bundler uyumlu `rubyzip` sürümleri bulamıyor.**
docxi 0.0.6, `rubyzip ~> 0.9.9` ile sabitlenmiştir ve bu, daha yeni bir rubyzip gerektiren her gem ile çakışır. docxi 0.1.0, rubyzip 1.3 ile 3.x arasını kabul eder.

**Bundler 0.0.6 kurmaya devam ediyor.**
Ya Gemfile docxi'yi bir 0.0.x sürümüne sabitliyordur, örneğin `"~> 0.0.6"` ile, ya da Ruby, 0.1.0'ın gerektirdiği 2.5'ten eskidir. Gemfile satırını `gem "docxi", "~> 0.1"` olarak değiştirin ve `ruby -v` çıktısını kontrol edin.

**nokogiri derlenemiyor.**
Platformunuza uyan önceden derlenmiş bir nokogiri gem'i yok; bu yüzden kaynaktan derleniyor. Bir C derleyicisi kurun ve [nokogiri kurulum rehberini](https://nokogiri.org/tutorials/installing_nokogiri.html) izleyin.

## <a id="errors-while-building-a-document"></a>Belge oluştururken çıkan hatalar

**Rails dışında `NoMethodError: undefined method 'present?'` hatası.**
docxi 0.0.6, ActiveSupport'un yüklü olduğunu varsayıyordu. 0.1.0 buna ihtiyaç duymaz.

**`tab`, `frame` veya bir üstbilginin `page_numbers` yöntemi çağrılırken `NameError` hatası.**
Bunlar 0.0.6'da bozuktu ve `document.frame("text")` dahil 0.1.0'da çalışır.

**Görsel eklerken `Errno::ENOENT` hatası.**
Yol, programınızın çalıştığı dizine göre mevcut değil. Ruby 3 üzerinde docxi 0.0.6 ile görsel URL'leri de bu şekilde başarısız olur, çünkü 0.0.6 `Kernel#open` kullanıyordu. 0.1.0, URL'leri `URI#open` ile indirir.

## <a id="documents-that-dont-look-right"></a>Doğru görünmeyen belgeler

**Google Docs veya LibreOffice dosyayı açmıyor.**
rubyzip 3 ile birlikte docxi 0.0.6 her girdiye Zip64 kayıtları yazar ve bu programlar onları reddeder. 0.1.0 arşivleri bu kayıtlar olmadan yazar.

**İlk sayfada üstbilgi veya altbilgi yok.**
docxi ilk sayfaya kendi üstbilgisini ve altbilgisini verir. Bunları `add_first_header` ve `add_first_footer` ile ayarlayın. [Yapılandırma](Configuration#first-page) sayfasına bakın.

**Bir kenar boşluğu eksik veya sayfa düzeni yanlış görünüyor.**
Kenar boşluklarını ayarlarken `top`, `right`, `bottom` ve `left` değerlerini birlikte verin. Atladıklarınız boş olarak yazılır.

**Metin parçaları arasındaki boşluklar kayboluyor.**
Word, bir metin parçasının başındaki veya sonundaki boşlukları siler. O `text` çağrısına `space: true` ekleyin.

**Üstbilgideki, tablo hücrelerindeki veya `pH`/`pV` kullanan görseller çok küçük ya da kare.**
0.1.0'da bilinen sorun: boyutları iki boyut için de `width * height * 14.92` EMU olarak hesaplanır. Boyutu `style:` ile verilen satır içi görseller kullanın veya [En iyi uygulamalar](Best-Practices#images) sayfasındaki geçici çözüme bakın.

**Aynı belge ikinci kez oluşturulduğunda bir görsel eksik oluyor.**
Açık bir `File` veya başka bir IO'dan eklenen görseller `render` sırasında okunur ve ikinci oluşturma boş bir akış okur. Bunun yerine bir yol verin ya da belgeyi bir kez oluşturun.

**`table` üzerinde ayarlanan kenarlıklar her hücrenin çevresinde görünüyor.**
Bilinen sorun: tablo düzeyindeki `borders:` tablonun dış kenarına değil, her hücreye uygulanır. Yalnızca dış çerçeveyi çizmek için `borders:` seçeneğini kenardaki hücrelerde, yalnızca dış kenarları için ayarlayın.

**Tablo hücrelerindeki sayfa numaraları "GlobalOptions" ile başlıyor.**
`TableCell#page_numbers` içinde bilinen bir sorundur. Sayfa numaralarını bunun yerine bir altbilgiye veya üstbilgiye koyun.

**Bir JPEG veya GIF görseli yanlış görünüyor ya da hiç görünmüyor.**
Her görsel, gerçek biçimi ne olursa olsun PNG içerik türüne sahip bir PNG parçası olarak saklanır; bu yüzden bazı programlar diğer biçimleri gösteremeyebilir. Görselleri eklemeden önce PNG'ye dönüştürün.

## <a id="inspecting-a-generated-file"></a>Üretilen bir dosyayı inceleme

Bir `.docx` dosyası zip arşividir; içine bakabilirsiniz:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Sorun devam ederse, sorunu yeniden üreten kısa bir betikle ve Ruby, docxi, nokogiri ve rubyzip sürümleriyle birlikte [bir issue açın](https://github.com/hammad417/docxi/issues).

## <a id="related-pages"></a>İlgili sayfalar

- [SSS](FAQ)
- [En iyi uygulamalar](Best-Practices)
- [Kurulum](Installation)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
