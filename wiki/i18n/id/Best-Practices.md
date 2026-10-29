<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Praktik terbaik docxi: menghindari kejutan umum

Kebiasaan-kebiasaan ini berasal langsung dari cara kerja docxi, dan sebagian besar mencegah masalah yang tercantum di [Pemecahan masalah](Troubleshooting). Semuanya dikelompokkan per topik. Saran pengaturan halaman berlaku untuk semua orang, sedangkan saran keamanan dan kinerja paling penting untuk aplikasi produksi.

## <a id="page-setup"></a>Pengaturan halaman

- **Berikan keempat margin sekaligus.** Hash yang Anda berikan ke `Docxi::Document.new` menggantikan nilai bawaan, dan nilai `top`, `right`, `bottom`, atau `left` mana pun yang tidak Anda isi akan ditulis kosong.
- **Atur header dan footer halaman pertama.** docxi selalu menyalakan halaman pertama yang terpisah. Jika halaman 1 memerlukan header atau footer, panggil juga `add_first_header` dan `add_first_footer`.

## <a id="text"></a>Teks

- **Gunakan `space: true` untuk spasi di awal atau akhir.** Jika tidak, Word membuangnya dan kata-kata dari dua run akan menempel.
- **Taruh format bersama di paragraf.** Opsi yang diberikan ke `p`, seperti `size:` atau `font:`, berlaku untuk setiap run di dalamnya, sehingga lebih sedikit yang perlu diulang.
- **Gunakan `p.br` untuk pemisah baris dan `document.br(page: true)` untuk pemisah halaman.** `document.br` tanpa `page: true` menghasilkan XML yang tidak valid.

## <a id="tables"></a>Tabel

- **Atur batas pada baris atau sel, bukan pada tabel.** Di 0.1.0, `borders:` tingkat tabel diterapkan ke setiap sel, bukan ke tepi luar tabel.
- **Berikan `columns_width` untuk setiap kolom.** Sel mengambil lebarnya dari situ, sehingga tabel mempertahankan tata letak yang Anda harapkan.

## <a id="images"></a>Gambar

- **Gunakan file PNG.** Setiap gambar disimpan sebagai part PNG.
- **Atur ukuran gambar sebaris dengan `style:`**, misalnya `style: { width: "120pt", height: "60pt" }`. Hanya cara ini yang menghasilkan ukuran gambar sesuai harapan di 0.1.0.
- **Berikan path, bukan `File` yang sudah dibuka, jika dokumen mungkin dirender lebih dari sekali.** Gambar IO dibaca selama `render`, dan render kedua mendapatkan gambar kosong.
- **Gambar header, gambar sel tabel, dan gambar berposisi sampai 0.2.0.** Lebar dan tingginya sama-sama dihitung sebagai `width * height * 14.92` EMU, sehingga hasilnya persegi. Untuk mendapatkan gambar persegi berukuran `S` inci, berikan `width` dan `height` yang sama dengan `Math.sqrt(S * 914400 / 14.92).round`, yaitu sekitar 248 untuk satu inci.

## <a id="security"></a>Keamanan

- **Jangan memberikan string yang tidak tepercaya ke `add_media`.** Path membuat docxi membaca file tersebut dari server Anda, dan URL membuat server Anda mengunduhnya. Berikan IO yang Anda buka sendiri, atau periksa nilainya terlebih dahulu. Lihat [Keamanan](Security).
- **Teks di-escape secara otomatis.** String yang diberikan ke `text` ditulis sebagai teks XML, sehingga input pengguna tidak bisa mengubah struktur dokumen.

## <a id="performance-and-reliability"></a>Kinerja dan keandalan

- **Buat dokumen besar di job latar belakang.** Seluruh dokumen disusun di memori saat Anda memanggil `render`. Lihat [Integrasi](Integrations).
- **Render sekali, lalu pakai ulang byte-nya.** Simpan `word.render.string` jika Anda perlu mengirim atau menyimpan dokumen lebih dari sekali.
- **Kunci versi yang kompatibel** dengan `gem "docxi", "~> 0.1"`.
- **Uji apa yang Anda hasilkan.** `.docx` adalah zip berisi file XML, jadi pengujian bisa mengekstraknya dan memeriksa XML-nya. Pengujian docxi sendiri bekerja dengan cara ini. Lihat [Arsitektur](Architecture).

## <a id="related-pages"></a>Halaman terkait

- [Fitur](Features)
- [Konfigurasi](Configuration)
- [Pemecahan masalah](Troubleshooting)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
