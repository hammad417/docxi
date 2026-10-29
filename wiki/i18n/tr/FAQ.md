<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi SSS: Ruby ve Rails'te Word belgeleri

En sık sorulan sorulara kısa yanıtlar. Her yanıtta daha ayrıntılı bir açıklamaya bağlantı vardır.

**Microsoft Word veya LibreOffice kurulu olmalı mı?**
Hayır. docxi XML'i ve zip arşivini kendisi yazar.

**docxi için Rails gerekli mi?**
Hayır. Her Ruby programında çalışır. Testler Rails 5.2 ile 8.1 arası yüklüyken de çalıştırılır. [Entegrasyonlar](Integrations) sayfasına bakın.

**Ruby on Rails'te Word belgesi nasıl oluşturulur?**
Belgeyi bir controller'da veya arka plan işinde oluşturun ve `word.render.string` değerini `send_data` ile gönderin. [Entegrasyonlar](Integrations) sayfasına bakın.

**Hangi Ruby ve Rails sürümleri destekleniyor?**
Ruby 2.5 ve daha yenisi. Rails isteğe bağlıdır ve CI, Rails 5.2 ile 8.1 arasını test eder. [Kurulum](Installation) sayfasına bakın.

**docxi var olan bir .docx dosyasını düzenleyebilir veya bir Word şablonunu doldurabilir mi?**
Hayır, yalnızca yeni belgeler oluşturur. Birleştirme alanları olan şablonlar için sablon gibi şablon tabanlı bir gem kullanın.

**İçindekiler tablosu başlıklardan mı oluşturuluyor?**
Hayır. Her girdiyi ve sayfa numarasını siz eklersiniz. [Özellikler](Features#table-of-contents) sayfasına bakın.

**Numaralı listeler veya başlık stilleri kullanabilir miyim?**
0.1.0'da hayır. Paragraflar madde işaretli öğe olabilir (`ul: true`) ve başlıkları `size:` ve `bold:` gibi metin seçenekleriyle biçimlendirebilirsiniz.

**Sayfa boyutunu veya yönünü değiştirebilir miyim?**
Henüz değil. Sayfalar dikey yönde US Letter boyutundadır ve yalnızca kenar boşluklarını değiştirebilirsiniz. [Yapılandırma](Configuration) sayfasına bakın.

**Hangi görsel biçimleri çalışıyor?**
PNG. Her görsel bir PNG parçası olarak saklanır.

**Bir URL'den görsel ekleyebilir miyim?**
Evet. `add_media`, `http`, `https` ve `ftp` URL'lerini kabul eder ve `render` çağrıldığında bunları indirir. Kullanıcılardan gelen URL'leri kontrol etmeden vermeyin. [Güvenlik](Security) sayfasına bakın.

**Belgeyi doğrudan bir dosyaya yazabilir miyim?**
`render` bir `StringIO` döndürür. Bunun `.string` değerini `File.binwrite("out.docx", word.render.string)` ile yazın.

**docxi çok büyük belgeleri akış (streaming) olarak üretebilir mi?**
Hayır. `render` çağrıldığında belgenin tamamı bellekte oluşturulur. [Dağıtım](Deployment) sayfasına bakın.

**Sağdan sola paragraflar ayarlayabilir miyim?**
0.1.0'da sağdan sola yazım için bir seçenek yoktur. Metin verdiğiniz gibi yazılır.

**Bilinen sorunlar nerede listeleniyor?**
[Yol haritası](Roadmap) sayfasında ve README'nin "Limitations and known issues" (sınırlamalar ve bilinen sorunlar) bölümünde.

## <a id="related-pages"></a>İlgili sayfalar

- [Sorun giderme](Troubleshooting)
- [Özellikler](Features)
- [Başlarken](Getting-Started)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
