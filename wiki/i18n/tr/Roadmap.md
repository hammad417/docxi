<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi yol haritası ve bilinen sorunlar

docxi'nin bir sürüm takvimi veya taahhüt edilmiş bir yol haritası yoktur. Kayıtlı tek planlı çalışma, aşağıdaki bilinen sorunları düzeltmektir. Bunlar tarih verilmeden 0.2.0 için planlanmıştır. Düzeltilmeleri mevcut kullanıcıların zaten ürettiği belgeleri değiştireceği için 0.1.0'a alınmadılar.

## <a id="known-issues-planned-for-020"></a>0.2.0 için planlanan bilinen sorunlar

| Sorun | 0.1.0'daki geçici çözüm |
|---|---|
| Konumlandırılmış görseller, üstbilgi görselleri ve tablo hücresi görselleri boyutlarını iki boyut için de `width * height * 14.92` EMU olarak hesaplar | Boyutu `style:` ile verilen satır içi görseller kullanın veya [En iyi uygulamalar](Best-Practices#images) sayfasına bakın |
| Tablo düzeyindeki `borders:` tablonun dış kenarına değil, her hücreye uygulanır | Yalnızca dış çerçeve için `borders:` seçeneğini kenardaki hücrelerin dış kenarlarında ayarlayın |
| Sayfa numarası alan kodunda `\* MERGEFORMAT` içindeki ters eğik çizgi eksik yazılır | Yok |
| Tablo hücrelerindeki sayfa numaralarının önüne "GlobalOptions" eklenir | Sayfa numaralarını bir üstbilgiye veya altbilgiye koyun |
| Görseller her zaman PNG parçası olarak saklanır ve `[Content_Types].xml` yalnızca ilk iki üstbilgi ile altbilgiyi bildirir | PNG görseller kullanın; en fazla iki üstbilgi ve iki altbilgi kullanın |
| Kenar boşluklarından yalnızca bazılarını vermek diğerlerini boş yazar | `top`, `right`, `bottom` ve `left` değerlerini birlikte verin |
| `docProps/core.xml` yerel saati UTC işaretiyle kaydeder ve yazarı hiç yazmaz | Yok |
| `Docxi.test` `NoMethodError` hatası verir | Çağırmayın. Eski sürümlerden kalmış bir yardımcıdır |
| `page: true` olmadan `document.br` ve `document.tab` geçersiz XML üretir | Bir paragrafın içinde `p.br` ve `p.tab` kullanın |

## <a id="not-planned"></a>Planlanmayanlar

Bunlar yukarıdaki listede yer almaz ve haklarında hiçbir karar verilmemiştir:

- var olan belgeleri düzenlemek veya şablonları doldurmak
- numaralı listeler, başlık stilleri veya otomatik içindekiler tablosu
- başka sayfa boyutları veya yatay sayfalar

## <a id="related-pages"></a>İlgili sayfalar

- [Değişiklik günlüğü](Changelog)
- [Katkıda bulunma](Contributing)
- [SSS](FAQ)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
