<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# docxi değişiklik günlüğü ve sürüm geçmişi

Bu sayfa her sürümü özetler. Notların tamamı depodaki [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) dosyasındadır; ikisi arasında fark olursa başvuru kaynağı odur.

## <a id="releases"></a>Sürümler

| Sürüm | Tarih | Notlar |
|---|---|---|
| 0.1.1 | 2026-10-29 | Gem özeti ve açıklaması düzeltildi, belgelere bağlantı eklendi. Kod değişikliği yok |
| 0.1.0 | 2026-10-28 | Güncel Ruby, Rails ve rubyzip ile çalışır. 0.0.6 ile aynı belgeleri üretir |
| 0.1.0.rc1 | 2026-10-28 | 0.1.0'ın ön sürümü |
| 0.0.6 | 2016-10-18 | 0.1.0 güncellemesinden önceki son sürüm |
| 0.0.2 | 2016-10-14 | Sürüm notu kaydedilmedi |
| 0.0.1 | 2016-10-14 | İlk sürüm. Sürüm notu kaydedilmedi |

0.0.3 ile 0.0.5 arasındaki sürümler RubyGems'te hiç yayımlanmadı.

## <a id="011"></a>0.1.1

Kod değişikliği yok ve üretilen belgeler 0.1.0 ile aynı.

- Gem özeti ve açıklaması artık docxi'nin ne yaptığını anlatıyor. Daha önce onu bir şablon oluşturucu olarak tanıtıyordu.
- RubyGems, gem belgeleri olarak bu wiki'ye bağlantı veriyor.

## <a id="010"></a>0.1.0

**Uyumluluk**

- Ruby 2.5 veya daha yenisini gerektirir. Daha eski Ruby kullanan uygulamalar 0.0.6'da kalır.
- Rails ile de Rails olmadan da çalışır ve Rails 5.2 ile 8.1 arasıyla birlikte test edilir.
- nokogiri için `>= 1.6.3, < 1.7` yerine `>= 1.8, < 2`; eski aralık 6.0'dan beri her Rails sürümüyle çakışıyordu.
- rubyzip için `~> 0.9.9` yerine `>= 1.3, < 4`.
- Artık ActiveSupport'un `blank?` ve `present?` yöntemlerini çağırmaz; bunlar yalnızca Rails onları yüklediğinde çalışıyordu.

**Düzeltmeler**

- rubyzip 3 ile yazılan belgeler artık Google Docs ve LibreOffice'in eski sürümlerinin reddettiği Zip64 kayıtlarını içermez.
- Görsel URL'si ile `add_media` Ruby 3'te yeniden çalışır.
- `add_media` artık `|` ile başlayan bir dizeyi kabuk komutu olarak çalıştırmaz.
- `tab` ve `frame` artık `NameError` hatası vermez.
- `document.frame("text")` metni içeren bir metin kutusu oluşturur.
- Bir üstbilgideki `page_numbers` artık `NameError` hatası vermez.

**Gem**

- Yeni bakımcı, ana sayfa ve değişiklik günlüğü bağlantıları.
- Sürüm yayımlamak MFA gerektirir.

## <a id="related-pages"></a>İlgili sayfalar

- [Yol haritası](Roadmap)
- [Kurulum](Installation#upgrading-from-00x)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
