<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# docxi'yi üretimde çalıştırma ve sürüm yayımlama

docxi bir servis değil, bir kütüphanedir; tek başına dağıtılacak bir şey yoktur. Bu sayfa, docxi kullanan bir uygulamayı yayına alırken nelere dikkat etmeniz gerektiğini ve docxi'nin yeni sürümlerinin nasıl yayımlandığını anlatır.

## <a id="shipping-an-app-that-uses-docxi"></a>docxi kullanan bir uygulamayı yayına almak

- **Bağımlılıklar.** `bundle install`, nokogiri ve rubyzip kurar. nokogiri, yaygın Linux, macOS ve Windows platformları için önceden derlenmiş gem'ler sunar. Diğer platformlarda, örneğin bazı minimal kapsayıcı imajlarında, kaynaktan derlenir ve bir C derleyicisi gerekir. [nokogiri kurulum rehberine](https://nokogiri.org/tutorials/installing_nokogiri.html) bakın.
- **Ofis yazılımı gerekmez.** Sunucular ve kapsayıcılar Microsoft Word veya LibreOffice gerektirmez.
- **Bellek.** `render`, belgenin tamamını bellekte oluşturur ve bir `StringIO` olarak döndürür. Bellek kullanımı sayfa sayısı ve görsellerin boyutu arttıkça artar.
- **Ağ ve dosyalar.** docxi dosya sistemini yalnızca `add_media` için verdiğiniz görsel yolları için okur, ağı da yalnızca verdiğiniz görsel URL'leri için kullanır. [Güvenlik](Security) sayfasına bakın.
- **Uzun belgeler.** Bunları bir web isteği sırasında değil, bir arka plan işinde oluşturun. [Entegrasyonlar](Integrations) sayfasına bakın.

## <a id="how-docxi-releases-are-published"></a>docxi sürümleri nasıl yayımlanır

Bu adımlar projenin bakımını yapanlar içindir. Sürümler, GitHub Actions üzerinden güvenilir yayınlama (trusted publishing) ile RubyGems'e gönderilir; bu yüzden depoda hiçbir API anahtarı saklanmaz.

1. `lib/docxi/version.rb` ve `CHANGELOG.md` dosyalarını güncelleyin ve `master` dalına birleştirin.
2. Commit'i başına `v` eklenmiş sürümle etiketleyin, örneğin `v0.1.1`, ve etiketi push edin.
3. `.github/workflows/release.yml`, gem'i derler ve RubyGems'e gönderir.

Ön sürümler de aynı adımları `0.2.0.rc1` gibi bir sürüm numarasıyla izler. Bundler ön sürümleri yalnızca açıkça istendiğinde kurar. CI (`.github/workflows/ci.yml`), `master` dalına yapılan her push'ta ve her pull request'te çalışır.

## <a id="related-pages"></a>İlgili sayfalar

- [Kurulum](Installation)
- [Geliştirme ortamı](Development-Setup)
- [Değişiklik günlüğü](Changelog)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
