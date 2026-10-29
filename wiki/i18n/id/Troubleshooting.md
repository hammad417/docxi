<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Pemecahan masalah docxi: galat instalasi, galat saat berjalan, dan keluaran yang rusak

Cari gejala Anda di bawah ini. Sebagian besar masalah di aplikasi lama berasal dari docxi 0.0.x, dan peningkatan ke 0.1.0 memperbaikinya tanpa perubahan kode.

## <a id="installation"></a>Instalasi

**Bundler tidak menemukan versi `nokogiri` yang kompatibel di aplikasi Rails 6+.**
docxi 0.0.6 mengunci `nokogiri < 1.7`, sedangkan setiap versi Rails sejak 6.0 memerlukan nokogiri 1.8.5 atau lebih baru. Gunakan `gem "docxi", "~> 0.1"`.

**Bundler tidak menemukan versi `rubyzip` yang kompatibel.**
docxi 0.0.6 mengunci `rubyzip ~> 0.9.9`, yang berbenturan dengan gem apa pun yang memerlukan rubyzip lebih baru. docxi 0.1.0 menerima rubyzip 1.3 sampai 3.x.

**Bundler terus memasang 0.0.6.**
Penyebabnya bisa Gemfile yang mengunci docxi ke versi 0.0.x, misalnya dengan `"~> 0.0.6"`, atau versi Ruby yang lebih lama dari 2.5, yang diperlukan 0.1.0. Ubah baris di Gemfile menjadi `gem "docxi", "~> 0.1"` dan periksa `ruby -v`.

**nokogiri gagal dikompilasi.**
Tidak ada gem nokogiri yang sudah dikompilasi untuk platform Anda, sehingga nokogiri dibangun dari kode sumber. Pasang kompiler C dan ikuti [panduan instalasi nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Galat saat membuat dokumen

**`NoMethodError: undefined method 'present?'` di luar Rails.**
docxi 0.0.6 menganggap ActiveSupport sudah dimuat. 0.1.0 tidak memerlukannya.

**`NameError` saat memanggil `tab`, `frame`, atau `page_numbers` milik header.**
Ketiganya rusak di 0.0.6 dan berfungsi di 0.1.0, termasuk `document.frame("text")`.

**`Errno::ENOENT` saat menambahkan gambar.**
Path tersebut tidak ada, relatif terhadap direktori tempat program Anda berjalan. Di Ruby 3 dengan docxi 0.0.6, URL gambar juga gagal dengan cara ini, karena 0.0.6 memakai `Kernel#open`. 0.1.0 mengunduh URL dengan `URI#open`.

## <a id="documents-that-dont-look-right"></a>Dokumen yang tampil tidak semestinya

**Google Docs atau LibreOffice tidak mau membuka file.**
docxi 0.0.6 dengan rubyzip 3 menulis rekaman Zip64 ke setiap entri, dan program-program ini menolaknya. 0.1.0 menulis arsip tanpa rekaman tersebut.

**Halaman pertama tidak punya header atau footer.**
docxi memberi halaman pertama header dan footer sendiri. Atur dengan `add_first_header` dan `add_first_footer`. Lihat [Konfigurasi](Configuration#first-page).

**Ada margin yang hilang, atau tata letak halaman terlihat salah.**
Saat mengatur margin, berikan `top`, `right`, `bottom`, dan `left` sekaligus. Yang tidak Anda isi akan ditulis kosong.

**Spasi di antara run teks menghilang.**
Word membuang spasi di awal atau akhir sebuah run. Tambahkan `space: true` ke panggilan `text` tersebut.

**Gambar di header, di sel tabel, atau dengan `pH`/`pV` sangat kecil atau berbentuk persegi.**
Masalah yang diketahui di 0.1.0: ukurannya dihitung sebagai `width * height * 14.92` EMU untuk kedua dimensi. Gunakan gambar sebaris yang ukurannya diatur dengan `style:`, atau lihat solusi sementara di [Praktik terbaik](Best-Practices#images).

**Gambar hilang saat dokumen yang sama dirender untuk kedua kalinya.**
Gambar yang ditambahkan dari `File` yang sudah dibuka atau IO lain dibaca selama `render`, dan render kedua membaca stream yang kosong. Berikan path sebagai gantinya, atau render sekali saja.

**Batas yang diatur pada `table` muncul di sekeliling setiap sel.**
Masalah yang diketahui: `borders:` tingkat tabel diterapkan ke setiap sel, bukan ke tepi luar tabel. Untuk menggambar garis luar saja, atur `borders:` pada sel-sel di tepi tabel, hanya untuk sisi luarnya.

**Nomor halaman di sel tabel diawali "GlobalOptions".**
Masalah yang diketahui pada `TableCell#page_numbers`. Taruh nomor halaman di footer atau header sebagai gantinya.

**Gambar JPEG atau GIF tampak salah atau tidak muncul.**
Setiap gambar disimpan sebagai part PNG dengan tipe konten PNG, apa pun format aslinya, sehingga beberapa program mungkin tidak menampilkan format lain. Konversikan gambar ke PNG sebelum menambahkannya.

## <a id="inspecting-a-generated-file"></a>Memeriksa file yang dihasilkan

File `.docx` adalah arsip zip, jadi Anda bisa melihat isinya:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Jika masalah masih ada, [buka issue](https://github.com/hammad417/docxi/issues) berisi skrip singkat yang mereproduksinya, beserta versi Ruby, docxi, nokogiri, dan rubyzip.

## <a id="related-pages"></a>Halaman terkait

- [Tanya jawab](FAQ)
- [Praktik terbaik](Best-Practices)
- [Instalasi](Installation)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
