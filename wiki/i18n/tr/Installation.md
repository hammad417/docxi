<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# docxi kurulumu: gereksinimler, Bundler ayarları ve yükseltme

docxi tamamen Ruby ile yazılmış bir gem'dir. Kurulduğunda yalnızca iki bağımlılık gelir: nokogiri ve rubyzip. Herhangi bir ofis yazılımı gerekmez.

## <a id="requirements"></a>Gereksinimler

| Gereksinim | Sürüm |
|---|---|
| Ruby | 2.5 veya daha yeni |
| nokogiri | `>= 1.8, < 2`, otomatik kurulur |
| rubyzip | `>= 1.3, < 4`, otomatik kurulur |
| Rails | Gerekli değil |

## <a id="install-with-bundler"></a>Bundler ile kurulum

Gem'i Gemfile dosyanıza ekleyin:

```ruby
gem "docxi", "~> 0.1"
```

Ardından kurun:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Bundler olmadan kurulum

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Kurulumu doğrulama

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Bundler projesinde bunu `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` olarak çalıştırın. Komut kurulu sürümü yazdırır, örneğin `0.1.0`.

## <a id="upgrading-from-00x"></a>0.0.x sürümünden yükseltme

Gemfile'ınız docxi'yi bir 0.0.x sürümüne sabitliyorsa, örneğin `"~> 0.0.6"` ile, satırı `gem "docxi", "~> 0.1"` olarak değiştirin. Ardından `bundle update docxi` komutunu çalıştırın. Kodda değişiklik gerekmez ve 0.1.0, 0.0.6 ile aynı belgeleri üretir: test paketi çıktıyı 0.0.6 tarafından üretilmiş dosyalarla karşılaştırır.

Arka planda neler değişti:

- nokogiri ve rubyzip artık yukarıdaki aralıklardaki herhangi bir sürüm olabilir; böylece docxi güncel Rails sürümlerinin yanına kurulabilir.
- docxi artık ActiveSupport'un yüklü olmasını gerektirmez.
- Arşivler Zip64 kayıtları olmadan yazılır.
- Görseller `Kernel#open` kullanılmadan yüklenir; bu sayede URL'ler Ruby 3'te çalışır.
- `tab`, `frame` ve üstbilgideki sayfa numaraları çalışır.

Ruby 2.5'ten eski bir sürümde çalışan uygulamalar 0.0.6 almaya devam eder, çünkü 0.1.0 en düşük sürüm olarak Ruby 2.5'i bildirir. [Değişiklik günlüğü](Changelog) tüm değişiklikleri listeler.

## <a id="versions-tested-in-ci"></a>CI'da test edilen sürümler

Her Rails sürümü, çatının tamamı docxi ile birlikte yüklenerek çalıştırılır.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI ayrıca şunları çalıştırır:

- Ruby 2.5, 2.6 ve 4.0 üzerinde Rails olmadan testler
- rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) ve nokogiri 1.11 (Ruby 2.5) ile testler
- LibreOffice'in örnek belgeleri dönüştürebildiğini doğrulayan bir kontrol

## <a id="platform-notes"></a>Platform notları

- nokogiri, yaygın Linux, macOS ve Windows platformları için önceden derlenmiş gem'ler sunar. Uygun olan yoksa kaynaktan derlenir ve bir C derleyicisi gerekir. [nokogiri kurulum rehberine](https://nokogiri.org/tutorials/installing_nokogiri.html) bakın.
- CI, Ubuntu 24.04 üzerinde çalışır ve testler macOS'te de çalıştırılmıştır. Windows test edilmez.

## <a id="if-installation-fails"></a>Kurulum başarısız olursa

nokogiri veya rubyzip üzerinde Bundler çakışmaları neredeyse her zaman eski docxi 0.0.6 sürümünün çözümlendiği anlamına gelir. [Sorun giderme](Troubleshooting#installation) sayfasına bakın.

## <a id="related-pages"></a>İlgili sayfalar

- [Başlarken](Getting-Started)
- [Yapılandırma](Configuration)
- [Dağıtım](Deployment)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
