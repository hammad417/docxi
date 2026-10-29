<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi güvenliği: sorun bildirme ve güvenilmeyen girdileri işleme

Bu sayfa, bir güvenlik sorununun nasıl bildirileceğini ve uygulamanız kullanıcı girdisini docxi'ye aktarırken bilmeniz gereken güvenlik noktalarını açıklar.

## <a id="reporting-a-vulnerability"></a>Güvenlik açığı bildirme

Lütfen güvenlik sorunlarını herkese açık issue'larda bildirmeyin. Bunun yerine projenin bakımını yapanlara e-posta gönderin. Adresler [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec) dosyasındadır. docxi sürümünü ve sorunu yeniden üretmenin kısa bir yolunu ekleyin.

## <a id="upgrade-from-00x"></a>0.0.x sürümünden yükseltin

docxi 0.0.x, görsel yolu dizelerini `Kernel#open` yöntemine aktarır. `Kernel#open`, `|` ile başlayan bir dizeyi kabuk komutu olarak çalıştırır; bu yüzden bir görsel yolunu kontrol eden bir saldırgan sunucunuzda komut çalıştırabilir. 0.1.0, dosyaları `File.open` ile açar ve URL'leri `URI#open` ile indirir. İkisi de komut çalıştırmaz.

## <a id="untrusted-input"></a>Güvenilmeyen girdiler

**Metin vermek güvenlidir.** Metin, tablo içerikleri ve diğer dizeler XML metni veya öznitelik değeri olarak yazılır ve nokogiri bunları kaçışlar. Kullanıcı girdisi belgeye öğe ekleyemez.

**Görsel kaynakları dikkat gerektirir.** `add_media` şunları kabul eder:

| Kaynak | docxi'nin yaptığı |
|---|---|
| Kendi açtığınız bir IO | Onu okur |
| Bir dosya yolu | O dosyayı sunucunuzun diskinden okur |
| Bir `http`, `https` veya `ftp` URL'si | Onu sunucunuzdan indirir |

Kullanıcı yolu seçebiliyorsa, sunucunuzdaki okunabilir herhangi bir dosyanın, örneğin yapılandırma dosyalarının veya anahtarların, belgeye gömülmesini sağlayabilir. Kullanıcı URL'yi seçebiliyorsa sunucunuz o adrese istek gönderir ve bu, iç adreslere erişebilir (sunucu taraflı istek sahteciliği, SSRF). Kendi açtığınız bir IO verin veya yalnızca bilinen yollara ve sunuculara izin verin.

**Köprüler** URL nereyi gösteriyorsa oraya gider. Kullanıcıların verdiği URL'leri `add_hyperlink` yöntemine vermeden önce kontrol edin.

## <a id="how-releases-are-protected"></a>Sürümler nasıl korunur

- Sürümler, GitHub Actions'tan RubyGems'in güvenilir yayınlama (trusted publishing) özelliğiyle yayımlanır; bu yüzden depoda hiçbir API anahtarı saklanmaz.
- Gem meta verileri `rubygems_mfa_required` ayarını içerir; bu yüzden kişisel bir API anahtarıyla yapılan gönderimler çok faktörlü kimlik doğrulama gerektirir.

## <a id="related-pages"></a>İlgili sayfalar

- [En iyi uygulamalar](Best-Practices)
- [Dağıtım](Deployment)
- [Katkıda bulunma](Contributing)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
