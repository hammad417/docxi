<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# Menjalankan docxi di produksi dan menerbitkan rilis

docxi adalah pustaka, bukan layanan, jadi tidak ada yang perlu di-deploy secara terpisah. Halaman ini membahas hal-hal penting saat Anda merilis aplikasi yang memakai docxi, serta cara versi baru docxi dirilis.

## <a id="shipping-an-app-that-uses-docxi"></a>Merilis aplikasi yang memakai docxi

- **Dependensi.** `bundle install` memasang nokogiri dan rubyzip. nokogiri menyediakan gem yang sudah dikompilasi untuk platform Linux, macOS, dan Windows yang umum. Di platform lain, misalnya beberapa image kontainer minimal, nokogiri dikompilasi dari kode sumber dan memerlukan kompiler C. Lihat [panduan instalasi nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Tanpa perangkat lunak perkantoran.** Server dan kontainer tidak memerlukan Microsoft Word atau LibreOffice.
- **Memori.** `render` menyusun seluruh dokumen di memori dan mengembalikannya sebagai `StringIO`. Penggunaan memori bertambah seiring jumlah halaman dan ukuran gambar.
- **Jaringan dan file.** docxi hanya membaca sistem file untuk path gambar yang Anda berikan ke `add_media`, dan hanya memakai jaringan untuk URL gambar yang Anda berikan. Lihat [Keamanan](Security).
- **Dokumen panjang.** Buat di job latar belakang, bukan selama permintaan web berlangsung. Lihat [Integrasi](Integrations).

## <a id="how-docxi-releases-are-published"></a>Cara rilis docxi diterbitkan

Langkah-langkah ini untuk pengelola proyek. Rilis dikirim ke RubyGems melalui GitHub Actions dengan trusted publishing, sehingga tidak ada kunci API yang disimpan di repositori.

1. Perbarui `lib/docxi/version.rb` dan `CHANGELOG.md`, lalu gabungkan ke `master`.
2. Beri tag pada commit dengan versi berawalan `v`, misalnya `v0.1.1`, lalu push tag tersebut.
3. `.github/workflows/release.yml` membangun gem dan mendorongnya ke RubyGems.

Pra-rilis mengikuti langkah yang sama dengan versi seperti `0.2.0.rc1`. Bundler hanya memasang pra-rilis jika diminta secara eksplisit. CI (`.github/workflows/ci.yml`) berjalan pada setiap push ke `master` dan setiap pull request.

## <a id="related-pages"></a>Halaman terkait

- [Instalasi](Installation)
- [Penyiapan pengembangan](Development-Setup)
- [Catatan perubahan](Changelog)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
