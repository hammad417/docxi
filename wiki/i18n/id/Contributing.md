<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Berkontribusi ke docxi

Laporan bug, perbaikan, dokumentasi, dan terjemahan semuanya disambut baik. Halaman ini menjelaskan cara melaporkan masalah dan apa yang diperlukan sebuah pull request. Untuk menyiapkan salinan kerja, lihat [Penyiapan pengembangan](Development-Setup).

## <a id="reporting-a-bug"></a>Melaporkan bug

Buka issue di [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) dengan menyertakan:

- skrip Ruby singkat yang mereproduksi masalahnya
- versi Ruby, docxi, nokogiri, dan rubyzip (serta Rails, jika Anda memakainya)
- apa yang Anda harapkan, dan apa yang Anda dapatkan. Untuk dokumen yang rusak, sebutkan program yang Anda pakai untuk membukanya

Laporkan masalah keamanan secara pribadi. Lihat [Keamanan](Security).

## <a id="pull-requests"></a>Pull request

1. Fork repositori dan buat branch.
2. Buat perubahannya, dan ikuti gaya kode di sekitarnya.
3. Tambahkan atau perbarui pengujian, lalu jalankan `bundle exec rake test`.
4. Jika XML yang dihasilkan sengaja berubah, jalankan `bundle exec rake golden` dan jelaskan perubahannya di pull request. File golden menjaga keluaran yang diandalkan pengguna saat ini.
5. Perbarui `CHANGELOG.md` dan dokumentasi jika perilakunya berubah.
6. Buka pull request. CI menjalankan pengujian di Ruby 2.5 sampai 4.0 dan Rails 5.2 sampai 8.1, serta memeriksa bahwa LibreOffice bisa membuka dokumen contoh.

Perubahan yang mengubah dokumen yang sudah ada, seperti perbaikan yang direncanakan untuk 0.2.0, sebaiknya dibuat dalam pull request tersendiri agar diff file golden mudah ditinjau.

## <a id="documentation-and-translations"></a>Dokumentasi dan terjemahan

- README bahasa Inggris adalah sumber dari 15 terjemahan di `docs/i18n/`. Ikuti `docs/i18n/TRANSLATION_GUIDE.md` di repositori saat Anda mengubahnya.
- Wiki ini ditulis di folder `wiki/` dalam repositori dan diterbitkan dari sana. `wiki/README.md` menjelaskan caranya.

## <a id="related-pages"></a>Halaman terkait

- [Penyiapan pengembangan](Development-Setup)
- [Arsitektur](Architecture)
- [Peta jalan](Roadmap)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
