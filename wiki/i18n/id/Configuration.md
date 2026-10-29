<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Mengonfigurasi docxi: margin halaman, pengaturan halaman, dan satuan

docxi tidak punya file konfigurasi, variabel lingkungan, atau pengaturan global. Semuanya diatur per dokumen atau per elemen. Halaman ini membahas pengaturan tingkat dokumen dan satuan yang dipakai oleh opsi. Opsi tiap elemen tercantum di [Fitur](Features).

## <a id="page-margins"></a>Margin halaman

Margin diberikan saat Anda membuat dokumen:

```ruby
require "docxi"

# Semua nilai dalam seperdua puluh poin (1440 = 1 inci).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Opsi | Bawaan | Arti |
|---|---|---|
| `top` | `720` (0,5 inci) | Margin atas |
| `right` | `1440` (1 inci) | Margin kanan |
| `bottom` | `720` (0,5 inci) | Margin bawah |
| `left` | `1440` (1 inci) | Margin kiri |
| `header` | `190` | Jarak dari tepi atas halaman ke header |
| `footer` | `200` | Jarak dari tepi bawah halaman ke footer |

Hash yang Anda berikan menggantikan nilai bawaan, jadi atur `top`, `right`, `bottom`, dan `left` sekaligus. Setiap nilai dari keempatnya yang tidak Anda isi akan ditulis sebagai nilai kosong. `header` dan `footer` kembali ke 190 dan 200 jika tidak ada atau kosong.

## <a id="page-size-and-orientation"></a>Ukuran dan orientasi halaman

Halaman berukuran US Letter dengan orientasi potret: 12240 × 15840 seperdua puluh poin, atau 8,5 × 11 inci. Ukuran dan orientasi tidak bisa diubah di 0.1.0.

## <a id="first-page"></a>Halaman pertama

docxi selalu menyalakan pengaturan Word "Halaman pertama berbeda". Header dan footer yang ditambahkan dengan `add_header` dan `add_footer` dipakai mulai halaman 2. Halaman 1 memakai `add_first_header` dan `add_first_footer`, jadi atur juga keduanya jika halaman pertama memerlukan header atau footer. Lihat [Fitur](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Isi setiap dokumen

- Font bawaan adalah Arial 12 pt.
- Gaya yang didefinisikan: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom`, dan `NoList`.
- Item daftar berpoin (`ul: true`) memakai tanda pisah sebagai poinnya.
- Properti dokumen mencatat nama aplikasi sebagai "Docxi Ruby Gem" dan waktu dokumen dirender.

## <a id="units"></a>Satuan

| Opsi | Satuan |
|---|---|
| `size:` pada teks | Poin (docxi menggandakannya menjadi setengah poin milik Word) |
| Margin, `left:` dan `right:` paragraf, `iwidth:` tabel, `height:` baris | Seperdua puluh poin (1440 per inci) |
| Ukuran `borders:` sel | Seperdelapan poin |
| `columns_width:` dan `width:` sel | Dikalikan 14.92 untuk mendapatkan seperdua puluh poin, mendekati piksel pada 96 dpi |
| `style:` gambar sebaris | Panjang CSS seperti `"120pt"` |
| `pH:` dan `pV:` gambar berposisi | EMU (914400 per inci) |
| Warna teks dan isian | Heksadesimal tanpa `#`, misalnya `"8B0000"` |
| `fill_color:` gambar | Warna VML dengan `#`, misalnya `"#00FF00"` |

## <a id="related-pages"></a>Halaman terkait

- [Fitur](Features)
- [Referensi API](API-Reference)
- [Praktik terbaik](Best-Practices)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
