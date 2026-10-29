<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Geliştirme ortamı: docxi testlerini ve araçlarını çalıştırma

Bu sayfa, docxi'nin bir çalışma kopyasını çalışır hale getirmenize yardım eder ve testleri, örnek belgeleri ve CI kurulumunu açıklar.

## <a id="get-the-code"></a>Kodu alın

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Geliştirme için 2.5 ve sonrasındaki her Ruby sürümü uygundur. Geliştirme bağımlılıkları minitest ve rake'tir.

## <a id="rake-tasks"></a>Rake görevleri

| Komut | Ne yapar |
|---|---|
| `bundle exec rake test` | Test paketini çalıştırır (varsayılan görev) |
| `bundle exec rake samples` | Örnek belgeleri Word, LibreOffice veya Google Docs ile açmak için `tmp/samples/*.docx` konumuna yazar |
| `bundle exec rake golden` | `test/golden` dizinini mevcut koddan yeniden yazar. Yalnızca çıktıda bilinçli bir değişiklikten sonra kullanın |
| `bundle exec rake build` | Gem'i `pkg/` dizinine derler |

## <a id="running-the-tests-with-rails-loaded"></a>Testleri Rails yüklüyken çalıştırma

`gemfiles/` klasöründe her Rails sürümü için bir Gemfile ve eski rubyzip ile nokogiri sürümleri için birkaç Gemfile daha bulunur:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1`, testlerden önce çatının tamamını yükler. Ayrıca docxi'nin `blank?` kurallarını ActiveSupport'unkilerle karşılaştırır.

## <a id="what-the-tests-cover"></a>Testlerin kapsadıkları

| Dosya | Kontrol ettiği |
|---|---|
| `test/test_golden_output.rb` | İki örnek belgenin her parçası 0.0.6'nın ürettiğiyle eşleşir |
| `test/test_package.rb` | XML iyi biçimlidir, ilişkiler var olan parçaları gösterir, her parçanın bir içerik türü vardır ve Zip64 kaydı yoktur |
| `test/test_media.rb` | Yollardan, IO nesnelerinden ve URL'lerden görseller ile `\|` dizesinin asla komut olarak çalıştırılmadığı |
| `test/test_fixed_features.rb` | `tab`, `frame` ve üstbilgideki sayfa numaraları |
| `test/test_util.rb` | `Docxi::Util.blank?` ActiveSupport kurallarına uyar |
| `test/test_rails_independence.rb` | docxi ActiveSupport olmadan çalışır |
| `test/test_warnings.rb` | Oluşturma sırasında `lib/` kaynaklı kullanımdan kaldırma uyarısı çıkmaz |

Örnek belgeler `test/support/sample_documents.rb` içindedir ve `test/support/package.rb` bir `.docx` dosyasını doğrudan zip yapılarından okur.

## <a id="continuous-integration"></a>Sürekli entegrasyon

`.github/workflows/ci.yml`, `master` dalına yapılan her push'ta ve her pull request'te çalışır. Kapsadıkları:

- Rails olmadan Ruby 2.5, 2.6 ve 4.0
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) ve 8.1 (4.0)
- başarısız olmasına izin verilen Rails `main`
- rubyzip 1.3 ve 2.4 ile nokogiri 1.11
- örnek belgeleri LibreOffice ile PDF'ye dönüştüren bir iş

Sürüm yayımlama [Dağıtım](Deployment) sayfasında anlatılır.

## <a id="related-pages"></a>İlgili sayfalar

- [Katkıda bulunma](Contributing)
- [Mimari](Architecture)
- [Dağıtım](Deployment)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
