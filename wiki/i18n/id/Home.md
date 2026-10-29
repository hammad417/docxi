<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki docxi: membuat dokumen Word (.docx) dengan Ruby

docxi adalah gem Ruby yang membuat file Microsoft Word `.docx` dari kode. Anda mendapatkan paragraf, teks berformat, tabel, gambar, header dan footer, nomor halaman, hyperlink, serta daftar isi sederhana. docxi tidak memerlukan Microsoft Word, LibreOffice, atau templat, dan berjalan di Ruby biasa maupun di aplikasi Ruby on Rails.

[README](https://github.com/hammad417/docxi#readme) adalah pengantar singkatnya. Wiki ini membahas setiap opsi, menyediakan resep yang siap disalin, menjelaskan cara kerja docxi di dalamnya, dan membantu ketika ada yang tidak beres.

Rilis saat ini: **0.1.1**, untuk Ruby 2.5 atau yang lebih baru. Lihat [Catatan perubahan](Changelog).

## <a id="start-here"></a>Mulai di sini

**Baru mengenal docxi**

- [Memulai](Getting-Started): pasang gem dan buat dokumen pertama Anda
- [Instalasi](Installation): persyaratan, Bundler, dan peningkatan dari 0.0.x
- [Contoh penggunaan](Usage-Examples): laporan, faktur, surat, dan tabel dari data, siap disalin

**Membuat dokumen**

- [Fitur](Features): setiap elemen beserta opsi yang diterimanya
- [Konfigurasi](Configuration): margin, pengaturan halaman, dan satuan
- [Referensi API](API-Reference): ringkasan kelas dan metode
- [Praktik terbaik](Best-Practices): kebiasaan yang menghindarkan kejutan umum

**Saat ada masalah**

- [Pemecahan masalah](Troubleshooting): pesan galat dan dokumen yang tampil tidak semestinya
- [Tanya jawab](FAQ): jawaban singkat untuk pertanyaan umum
- [Keamanan](Security): melaporkan masalah dan menangani input yang tidak tepercaya

**Lebih dalam**

- [Arsitektur](Architecture): bagaimana panggilan Ruby menjadi paket `.docx`
- [Integrasi](Integrations): controller Rails, aplikasi Rack, dan job latar belakang
- [Deployment](Deployment): menjalankan docxi di produksi, dan cara rilis diterbitkan
- [Berkontribusi](Contributing) dan [Penyiapan pengembangan](Development-Setup)
- [Peta jalan](Roadmap), [Catatan perubahan](Changelog), dan [Kredit](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>Apa yang dilakukan docxi, dan apa yang tidak

docxi membuat dokumen baru. docxi tidak bisa membuka atau mengedit `.docx` yang sudah ada, mengisi templat Word, atau menjalankan mail merge. Tidak ada daftar bernomor, gaya judul, atau daftar isi otomatis, dan halaman selalu berukuran US Letter dengan orientasi potret. Masalah yang diketahui dan direncanakan untuk 0.2.0 tercantum di [Peta jalan](Roadmap).

## <a id="credits"></a>Kredit

docxi dikelola oleh [Hammad Maqbool](https://github.com/hammad417) di [Phaedra Solutions](https://www.phaedrasolutions.com/). Proyek ini dimulai pada 2016 dari [docxer](https://github.com/gtimti/docxer) karya Timur Gabdrakipov. Lihat [Kredit](Credits) untuk detailnya.

## <a id="languages"></a>Bahasa

Versi bahasa Inggris adalah versi acuan wiki ini. Tautan ke terjemahan ada di bilah samping dan di bagian atas setiap halaman.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Laporkan masalah](https://github.com/hammad417/docxi/issues)
