<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# docxi en iyi uygulamaları: yaygın sürprizlerden kaçının

Bu alışkanlıklar doğrudan docxi'nin çalışma biçiminden gelir ve çoğu, [Sorun giderme](Troubleshooting) sayfasında listelenen bir sorunu önler. Konulara göre gruplandırılmıştır. Sayfa yapısıyla ilgili öneriler herkes için geçerlidir; güvenlik ve performans önerileri ise en çok üretimdeki uygulamalar için önemlidir.

## <a id="page-setup"></a>Sayfa yapısı

- **Dört kenar boşluğunu birlikte verin.** `Docxi::Document.new` için verdiğiniz hash varsayılanların yerine geçer ve `top`, `right`, `bottom` ile `left` değerlerinden atladığınız her biri boş yazılır.
- **İlk sayfanın üstbilgisini ve altbilgisini ayarlayın.** docxi her zaman ayrı bir ilk sayfa açar. 1. sayfada üstbilgi veya altbilgi gerekiyorsa `add_first_header` ve `add_first_footer` yöntemlerini de çağırın.

## <a id="text"></a>Metin

- **Baştaki veya sondaki boşluklar için `space: true` kullanın.** Aksi halde Word bunları siler ve iki metin parçasının kelimeleri birbirine değer.
- **Ortak biçimlendirmeyi paragrafa koyun.** `p` için verilen `size:` veya `font:` gibi seçenekler içindeki her metin parçasına uygulanır; böylece daha az tekrar edersiniz.
- **Satır sonları için `p.br`, sayfa sonları için `document.br(page: true)` kullanın.** `page: true` olmadan `document.br` geçersiz XML üretir.

## <a id="tables"></a>Tablolar

- **Kenarlıkları tabloda değil, satırlarda veya hücrelerde ayarlayın.** 0.1.0'da tablo düzeyindeki `borders:` tablonun dış kenarına değil, her hücreye uygulanır.
- **Her sütun için `columns_width` verin.** Hücreler genişliklerini buradan alır; böylece tablolar beklediğiniz düzeni korur.

## <a id="images"></a>Görseller

- **PNG dosyaları kullanın.** Her görsel bir PNG parçası olarak saklanır.
- **Satır içi görsellerin boyutunu `style:` ile verin**, örneğin `style: { width: "120pt", height: "60pt" }`. 0.1.0'da beklendiği gibi çalışan tek görsel boyutu budur.
- **Belge birden fazla kez oluşturulabilecekse açık bir `File` yerine bir yol verin.** IO görselleri `render` sırasında okunur ve ikinci oluşturma boş bir görsel alır.
- **0.2.0'a kadar üstbilgi, tablo hücresi ve konumlandırılmış görseller.** Genişlik ve yüksekliklerinin ikisi de `width * height * 14.92` EMU olarak hesaplanır; bu yüzden kare çıkarlar. Kenarı `S` inç olan kare bir görsel için `width` ve `height` değerlerini `Math.sqrt(S * 914400 / 14.92).round` olarak verin; bir inç için bu yaklaşık 248'dir.

## <a id="security"></a>Güvenlik

- **`add_media` yöntemine güvenilmeyen dizeler vermeyin.** Yol verildiğinde docxi o dosyayı sunucunuzdan okur, URL verildiğinde de sunucunuz onu indirir. Kendi açtığınız bir IO verin veya önce değeri kontrol edin. [Güvenlik](Security) sayfasına bakın.
- **Metin sizin için kaçışlanır.** `text` için verilen dizeler XML metni olarak yazılır; bu yüzden kullanıcı girdisi belgenin yapısını değiştiremez.

## <a id="performance-and-reliability"></a>Performans ve güvenilirlik

- **Büyük belgeleri bir arka plan işinde oluşturun.** `render` çağrıldığında belgenin tamamı bellekte oluşturulur. [Entegrasyonlar](Integrations) sayfasına bakın.
- **Bir kez oluşturun, sonra baytları yeniden kullanın.** Belgeyi birden fazla kez göndermeniz veya saklamanız gerekiyorsa `word.render.string` değerini saklayın.
- **Uyumlu bir sürümü sabitleyin**: `gem "docxi", "~> 0.1"` ile.
- **Ürettiğiniz şeyi test edin.** Bir `.docx` dosyası XML dosyalarından oluşan bir zip'tir; testler onu açıp XML'i kontrol edebilir. docxi'nin kendi testleri de böyle çalışır. [Mimari](Architecture) sayfasına bakın.

## <a id="related-pages"></a>İlgili sayfalar

- [Özellikler](Features)
- [Yapılandırma](Configuration)
- [Sorun giderme](Troubleshooting)

---

[← Wiki ana sayfası](Home) · [README](https://github.com/hammad417/docxi#readme)
