<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Penyiapan pengembangan: menjalankan pengujian dan alat docxi

Halaman ini membantu Anda menjalankan salinan kerja docxi, serta menjelaskan pengujian, dokumen contoh, dan penyiapan CI.

## <a id="get-the-code"></a>Ambil kodenya

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Ruby versi apa pun mulai 2.5 bisa dipakai untuk pengembangan. Dependensi pengembangannya adalah minitest dan rake.

## <a id="rake-tasks"></a>Tugas Rake

| Perintah | Fungsinya |
|---|---|
| `bundle exec rake test` | Menjalankan rangkaian pengujian (tugas bawaan) |
| `bundle exec rake samples` | Menulis dokumen contoh ke `tmp/samples/*.docx`, untuk dibuka di Word, LibreOffice, atau Google Docs |
| `bundle exec rake golden` | Menulis ulang `test/golden` dari kode saat ini. Gunakan hanya setelah perubahan keluaran yang disengaja |
| `bundle exec rake build` | Membangun gem ke `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Menjalankan pengujian dengan Rails termuat

Folder `gemfiles/` berisi satu Gemfile per versi Rails, ditambah beberapa untuk versi rubyzip dan nokogiri yang lebih lama:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` memuat seluruh framework sebelum pengujian. Variabel ini juga membandingkan aturan `blank?` milik docxi dengan milik ActiveSupport.

## <a id="what-the-tests-cover"></a>Cakupan pengujian

| File | Yang diperiksa |
|---|---|
| `test/test_golden_output.rb` | Setiap part dari dua dokumen contoh sama dengan yang dihasilkan 0.0.6 |
| `test/test_package.rb` | XML-nya well-formed, relasi menunjuk ke part yang ada, setiap part punya tipe konten, dan tidak ada rekaman Zip64 |
| `test/test_media.rb` | Gambar dari path, objek IO, dan URL, serta bahwa string `\|` tidak pernah dijalankan sebagai perintah |
| `test/test_fixed_features.rb` | `tab`, `frame`, dan nomor halaman di header |
| `test/test_util.rb` | `Docxi::Util.blank?` mengikuti aturan ActiveSupport |
| `test/test_rails_independence.rb` | docxi berjalan tanpa ActiveSupport |
| `test/test_warnings.rb` | Proses render tidak memunculkan peringatan deprecation dari `lib/` |

Dokumen contoh ada di `test/support/sample_documents.rb`, dan `test/support/package.rb` membaca `.docx` langsung dari struktur zip-nya.

## <a id="continuous-integration"></a>Integrasi berkelanjutan

`.github/workflows/ci.yml` berjalan pada setiap push ke `master` dan setiap pull request. Cakupannya:

- Ruby 2.5, 2.6, dan 4.0 tanpa Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4), dan 8.1 (4.0)
- Rails `main`, yang boleh gagal
- rubyzip 1.3 dan 2.4, serta nokogiri 1.11
- job yang mengonversi dokumen contoh ke PDF dengan LibreOffice

Rilis dibahas di [Deployment](Deployment).

## <a id="related-pages"></a>Halaman terkait

- [Berkontribusi](Contributing)
- [Arsitektur](Architecture)
- [Deployment](Deployment)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
