<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# docxi'ye katkıda bulunma

Hata bildirimleri, düzeltmeler, belgeler ve çeviriler memnuniyetle karşılanır. Bu sayfa, bir sorunun nasıl bildirileceğini ve bir pull request'in nelere ihtiyaç duyduğunu açıklar. Çalışma kopyası hazırlamak için [Geliştirme ortamı](Development-Setup) sayfasına bakın.

## <a id="reporting-a-bug"></a>Hata bildirme

[github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) adresinde şunları içeren bir issue açın:

- sorunu yeniden üreten kısa bir Ruby betiği
- Ruby, docxi, nokogiri ve rubyzip sürümleri (kullanıyorsanız Rails sürümü de)
- ne beklediğiniz ve ne elde ettiğiniz. Bozuk belgeler için onları hangi programda açtığınızı belirtin

Güvenlik sorunlarını bunun yerine gizli olarak bildirin. [Güvenlik](Security) sayfasına bakın.

## <a id="pull-requests"></a>Pull request'ler

1. Depoyu fork edin ve bir dal oluşturun.
2. Değişikliği yapın ve çevresindeki kodun stiline uyun.
3. Testleri ekleyin veya güncelleyin ve `bundle exec rake test` komutunu çalıştırın.
4. Üretilen XML bilerek değişiyorsa `bundle exec rake golden` komutunu çalıştırın ve değişikliği pull request'te açıklayın. Golden dosyalar, mevcut kullanıcıların dayandığı çıktıyı korur.
5. Davranış değişiyorsa `CHANGELOG.md` dosyasını ve belgeleri güncelleyin.
6. Pull request'i açın. CI, testleri Ruby 2.5 ile 4.0 ve Rails 5.2 ile 8.1 arasında çalıştırır ve LibreOffice'in örnek belgeleri açabildiğini kontrol eder.

0.2.0 için planlanan düzeltmeler gibi mevcut belgeleri değiştiren değişiklikler, golden dosya farkının kolayca incelenebilmesi için ayrı bir pull request'te yer almalıdır.

## <a id="documentation-and-translations"></a>Belgeler ve çeviriler

- İngilizce README, `docs/i18n/` içindeki 15 çevirinin kaynağıdır. Onu değiştirirken depodaki `docs/i18n/TRANSLATION_GUIDE.md` dosyasını izleyin.
- Bu wiki, deponun `wiki/` klasöründe yazılır ve oradan yayımlanır. Nasıl yapıldığını `wiki/README.md` açıklar.

## <a id="related-pages"></a>İlgili sayfalar

- [Geliştirme ortamı](Development-Setup)
- [Mimari](Architecture)
- [Yol haritası](Roadmap)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
