<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi wiki: Ruby ile Word (.docx) belgeleri oluşturun

docxi, koddan Microsoft Word `.docx` dosyaları oluşturan bir Ruby gem'idir. Paragraflar, biçimlendirilmiş metin, tablolar, görseller, üstbilgi ve altbilgiler, sayfa numaraları, köprüler ve basit bir içindekiler tablosu sunar. Microsoft Word, LibreOffice ya da bir şablon gerektirmez; düz Ruby'de de Ruby on Rails uygulamalarında da çalışır.

[README](https://github.com/hammad417/docxi#readme) kısa bir giriştir. Bu wiki her seçeneği ele alır, kopyalanmaya hazır tarifler sunar, docxi'nin içeride nasıl çalıştığını açıklar ve bir şeyler ters gittiğinde yardımcı olur.

Güncel sürüm: **0.1.1**, Ruby 2.5 veya daha yenisi için. [Değişiklik günlüğüne](Changelog) bakın.

## <a id="start-here"></a>Buradan başlayın

**docxi'yi ilk kez kullanıyorsanız**

- [Başlarken](Getting-Started): gem'i kurun ve ilk belgenizi oluşturun
- [Kurulum](Installation): gereksinimler, Bundler ve 0.0.x sürümünden yükseltme
- [Kullanım örnekleri](Usage-Examples): verilerden raporlar, faturalar, mektuplar ve tablolar, kopyalanmaya hazır

**Belge oluşturma**

- [Özellikler](Features): her öğe ve kabul ettiği seçenekler
- [Yapılandırma](Configuration): kenar boşlukları, sayfa yapısı ve birimler
- [API başvurusu](API-Reference): sınıflar ve yöntemlere hızlı bakış
- [En iyi uygulamalar](Best-Practices): yaygın sürprizlerden kaçındıran alışkanlıklar

**Bir şeyler ters gittiğinde**

- [Sorun giderme](Troubleshooting): hata mesajları ve doğru görünmeyen belgeler
- [SSS](FAQ): sık sorulan sorulara kısa yanıtlar
- [Güvenlik](Security): sorun bildirme ve güvenilmeyen girdileri işleme

**Daha fazlası**

- [Mimari](Architecture): Ruby çağrılarının nasıl `.docx` paketine dönüştüğü
- [Entegrasyonlar](Integrations): Rails controller'ları, Rack uygulamaları ve arka plan işleri
- [Dağıtım](Deployment): docxi'yi üretimde çalıştırmak ve sürümlerin nasıl yayımlandığı
- [Katkıda bulunma](Contributing) ve [Geliştirme ortamı](Development-Setup)
- [Yol haritası](Roadmap), [Değişiklik günlüğü](Changelog) ve [Emeği geçenler](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi neler yapar, neler yapmaz

docxi yeni belgeler oluşturur. Var olan bir `.docx` dosyasını açamaz veya düzenleyemez, bir Word şablonunu dolduramaz ve adres mektup birleştirme yapamaz. Numaralı listeler, başlık stilleri ve otomatik içindekiler tablosu yoktur; sayfalar her zaman dikey yönde US Letter boyutundadır. 0.2.0 için planlanan bilinen sorunlar [Yol haritası](Roadmap) sayfasında listelenir.

## <a id="credits"></a>Emeği geçenler

docxi'nin bakımını [Phaedra Solutions](https://www.phaedrasolutions.com/) bünyesinde [Hammad Maqbool](https://github.com/hammad417) yapar. Proje 2016'da Timur Gabdrakipov'un [docxer](https://github.com/gtimti/docxer) projesinden yola çıkarak başladı. Ayrıntılar için [Emeği geçenler](Credits) sayfasına bakın.

## <a id="languages"></a>Diller

Bu wiki'nin başvuru sürümü İngilizce sürümdür. Çevirilerin bağlantıları kenar çubuğunda ve her sayfanın başında yer alır.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Sorun bildirin](https://github.com/hammad417/docxi/issues)
