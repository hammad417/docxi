<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# Tanya jawab docxi: dokumen Word di Ruby dan Rails

Jawaban singkat untuk pertanyaan yang paling sering diajukan. Setiap jawaban menautkan ke penjelasan yang lebih panjang.

**Apakah Microsoft Word atau LibreOffice harus terpasang?**
Tidak. docxi menulis XML dan arsip zip-nya sendiri.

**Apakah docxi memerlukan Rails?**
Tidak. docxi berjalan di program Ruby apa pun. Pengujiannya juga dijalankan dengan Rails 5.2 sampai 8.1 dalam keadaan termuat. Lihat [Integrasi](Integrations).

**Bagaimana cara membuat dokumen Word di Ruby on Rails?**
Susun dokumennya di controller atau job latar belakang, lalu kirim `word.render.string` dengan `send_data`. Lihat [Integrasi](Integrations).

**Versi Ruby dan Rails mana yang didukung?**
Ruby 2.5 dan yang lebih baru. Rails bersifat opsional, dan CI menguji Rails 5.2 sampai 8.1. Lihat [Instalasi](Installation).

**Bisakah docxi mengedit .docx yang sudah ada atau mengisi templat Word?**
Tidak, docxi hanya membuat dokumen baru. Untuk templat dengan bidang penggabungan, gunakan gem berbasis templat seperti sablon.

**Apakah daftar isi dibuat dari judul?**
Tidak. Anda sendiri yang menambahkan setiap entri beserta nomor halamannya. Lihat [Fitur](Features#table-of-contents).

**Bisakah saya memakai daftar bernomor atau gaya judul?**
Tidak di 0.1.0. Paragraf bisa dijadikan item berpoin (`ul: true`), dan Anda bisa memformat judul dengan opsi teks seperti `size:` dan `bold:`.

**Bisakah saya mengubah ukuran atau orientasi halaman?**
Belum. Halaman berukuran US Letter dengan orientasi potret, dan Anda hanya bisa mengubah margin. Lihat [Konfigurasi](Configuration).

**Format gambar apa yang berfungsi?**
PNG. Setiap gambar disimpan sebagai part PNG.

**Bisakah saya menambahkan gambar dari URL?**
Bisa. `add_media` menerima URL `http`, `https`, dan `ftp`, lalu mengunduhnya saat Anda memanggil `render`. Jangan memberikan URL dari pengguna tanpa memeriksanya. Lihat [Keamanan](Security).

**Bisakah saya menulis dokumen langsung ke file?**
`render` mengembalikan `StringIO`. Tulis `.string` miliknya dengan `File.binwrite("out.docx", word.render.string)`.

**Bisakah docxi melakukan streaming untuk dokumen yang sangat besar?**
Tidak. Seluruh dokumen disusun di memori saat Anda memanggil `render`. Lihat [Deployment](Deployment).

**Bisakah saya mengatur paragraf kanan-ke-kiri?**
Tidak ada opsi kanan-ke-kiri di 0.1.0. Teks ditulis persis seperti yang Anda berikan.

**Di mana daftar masalah yang diketahui?**
Di [Peta jalan](Roadmap) dan di bagian "Limitations and known issues" (batasan dan masalah yang diketahui) pada README.

## <a id="related-pages"></a>Halaman terkait

- [Pemecahan masalah](Troubleshooting)
- [Fitur](Features)
- [Memulai](Getting-Started)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
