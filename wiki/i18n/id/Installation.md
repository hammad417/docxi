<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Menginstal docxi: persyaratan, penyiapan Bundler, dan peningkatan

docxi adalah gem yang sepenuhnya ditulis dengan Ruby. Pemasangannya hanya menarik dua dependensi, nokogiri dan rubyzip, tidak lebih. Tidak diperlukan perangkat lunak perkantoran.

## <a id="requirements"></a>Persyaratan

| Persyaratan | Versi |
|---|---|
| Ruby | 2.5 atau lebih baru |
| nokogiri | `>= 1.8, < 2`, dipasang otomatis |
| rubyzip | `>= 1.3, < 4`, dipasang otomatis |
| Rails | Tidak diperlukan |

## <a id="install-with-bundler"></a>Instal dengan Bundler

Tambahkan gem ke Gemfile Anda:

```ruby
gem "docxi", "~> 0.1"
```

Lalu pasang:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Instal tanpa Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Periksa instalasi

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Di proyek Bundler, jalankan sebagai `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. Perintah ini mencetak versi yang terpasang, misalnya `0.1.0`.

## <a id="upgrading-from-00x"></a>Meningkatkan dari 0.0.x

Jika Gemfile Anda mengunci docxi ke versi 0.0.x, misalnya dengan `"~> 0.0.6"`, ubah barisnya menjadi `gem "docxi", "~> 0.1"`. Lalu jalankan `bundle update docxi`. Kode tidak perlu diubah, dan 0.1.0 menghasilkan dokumen yang sama dengan 0.0.6: rangkaian pengujian membandingkan keluarannya dengan file yang dihasilkan oleh 0.0.6.

Yang berubah di balik layar:

- nokogiri dan rubyzip kini boleh berversi apa saja dalam rentang di atas, sehingga docxi bisa dipasang berdampingan dengan versi Rails terkini.
- docxi tidak lagi memerlukan ActiveSupport dalam keadaan termuat.
- Arsip ditulis tanpa rekaman Zip64.
- Gambar dimuat tanpa `Kernel#open`, sehingga URL berfungsi di Ruby 3.
- `tab`, `frame`, dan nomor halaman di header berfungsi.

Aplikasi yang memakai Ruby lebih lama dari 2.5 tetap mendapatkan 0.0.6, karena 0.1.0 menyatakan Ruby 2.5 sebagai versi minimum. [Catatan perubahan](Changelog) mencantumkan setiap perubahan.

## <a id="versions-tested-in-ci"></a>Versi yang diuji di CI

Setiap versi Rails dijalankan dengan seluruh framework dimuat bersama docxi.

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

CI juga menjalankan:

- pengujian tanpa Rails di Ruby 2.5, 2.6, dan 4.0
- pengujian dengan rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3), dan nokogiri 1.11 (Ruby 2.5)
- pemeriksaan bahwa LibreOffice bisa mengonversi dokumen contoh

## <a id="platform-notes"></a>Catatan platform

- nokogiri menyediakan gem yang sudah dikompilasi untuk platform Linux, macOS, dan Windows yang umum. Jika tidak ada yang cocok, nokogiri dikompilasi dari kode sumber dan memerlukan kompiler C. Lihat [panduan instalasi nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- CI berjalan di Ubuntu 24.04, dan pengujiannya juga pernah dijalankan di macOS. Windows tidak diuji.

## <a id="if-installation-fails"></a>Jika instalasi gagal

Konflik Bundler pada nokogiri atau rubyzip hampir selalu berarti docxi 0.0.6 versi lama yang sedang diresolusi. Lihat [Pemecahan masalah](Troubleshooting#installation).

## <a id="related-pages"></a>Halaman terkait

- [Memulai](Getting-Started)
- [Konfigurasi](Configuration)
- [Deployment](Deployment)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
