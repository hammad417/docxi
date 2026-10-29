<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Fitur docxi: paragraf, teks, tabel, gambar, header, dan footer

Halaman ini mencantumkan setiap elemen yang bisa ditambahkan docxi ke dokumen Word beserta opsi yang diterima masing-masing. Contoh-contohnya dibangun di atas penyiapan berikut:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Opsi yang tidak Anda perlukan boleh dihilangkan. Warna berupa nilai heksadesimal tanpa `#`, misalnya `"8B0000"`. Untuk satuan dan pengaturan halaman, lihat [Konfigurasi](Configuration).

## <a id="paragraphs"></a>Paragraf

`document.p(options) { |p| ... }` menambahkan paragraf.

| Opsi | Nilai | Efek |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Perataan. `"both"` meratakan teks kiri-kanan |
| `fill` | Warna heksadesimal | Arsiran latar di belakang paragraf dan teksnya |
| `left`, `right` | Seperdua puluh poin | Indentasi. Atur keduanya, karena jika hanya satu, nilainya diabaikan |
| `ul` | `true` | Menjadikan paragraf sebagai item daftar berpoin |
| `bottom` | `true` | Menggambar batas di bawah paragraf |

Opsi teks seperti `size:`, `bold:`, atau `color:` juga bisa diberikan ke `p`. Opsi itu menjadi nilai bawaan untuk setiap run teks di paragraf tersebut.

Di dalam blok, Anda bisa memanggil `text`, `br`, `tab`, `image`, `hyperlink`, dan `frame`. `document.text(string, options)` adalah jalan pintas untuk paragraf yang hanya berisi satu run.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Teks

`p.text(string, options)` menambahkan satu run teks.

| Opsi | Nilai | Efek |
|---|---|---|
| `bold` | `true` | Tebal |
| `italic` | `true` | Miring |
| `underline` | `true` | Garis bawah tunggal |
| `color` | Warna heksadesimal | Warna teks |
| `size` | Poin | Ukuran font |
| `font` | Nama font | Font, misalnya `"Georgia"` |
| `fill` | Warna heksadesimal | Arsiran latar di belakang teks |
| `space` | `true` | Mempertahankan spasi di awal dan akhir teks |
| `br` | `true` | Menambahkan pemisah baris setelah teks |

String `nil` dilewati.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Pemisah baris, pemisah halaman, dan tab

- `p.br`, atau `p.br(times: 3)`, menambahkan pemisah baris di dalam paragraf.
- `document.br(page: true)` memulai halaman baru.
- `p.tab`, atau `p.tab(times: 2)`, menambahkan karakter tab di dalam paragraf.
- `document.tab(page: true)` menambahkan paragraf yang hanya berisi satu tab.

Jangan memanggil `document.br` atau `document.tab` tanpa `page: true`. Keduanya menaruh run teks langsung di isi dokumen, dan itu bukan WordprocessingML yang valid.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tabel

`document.table(options)` menerima blok. Di dalamnya, `table.tr(options)` menambahkan baris dan `tr.tc(options)` menambahkan sel.

**Opsi tabel**

| Opsi | Nilai | Efek |
|---|---|---|
| `columns_width` | Array angka | Lebar kolom, dikalikan 14.92 menjadi seperdua puluh poin. Juga menjadi lebar bawaan sel |
| `height` | Seperdua puluh poin | Tinggi minimum setiap baris |
| `iwidth` | Seperdua puluh poin | Indentasi tabel dari margin kiri |
| `style` | ID gaya | Gaya tabel. docxi mendefinisikan `TableNormal` dan `TableCustom` |
| `width` | Angka | Lebar tabel yang diinginkan, ditulis dengan tipe `auto`. Tanpa `columns_width`, setiap sel juga mendapat lebar ini, dikalikan 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Perataan tabel. Teks yang ditambahkan dengan `tc.text` diratakan dengan cara yang sama |
| `padding` | `true` | Menambahkan padding tetap di dalam sel |
| `borders` | Hash | Masalah yang diketahui: diterapkan ke setiap sel, bukan ke tepi luar tabel |

**Opsi baris**: `height:` mengatur tinggi minimum baris. Setiap opsi sel yang diberikan ke `tr` berlaku untuk setiap sel di baris tersebut. Opsi tabel diteruskan dengan cara yang sama ke setiap baris dan sel.

**Opsi sel**

| Opsi | Nilai | Efek |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Lebar batas dalam seperdelapan poin. `nil` berarti tidak ada batas di sisi itu |
| `fill` | Warna heksadesimal | Arsiran latar sel. Sel tanpa isian berwarna putih (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Perataan vertikal |
| `colspan` | Bilangan bulat | Jumlah kolom yang direntang sel |
| `rowspan` | `true` | Memulai penggabungan vertikal |
| `merged` | `true` | Melanjutkan penggabungan vertikal dari sel di atasnya. Isinya diabaikan |
| `width` | Angka | Lebar sel, dikalikan 14.92 |

**Isi sel**

- `tc.text(string, options)`: opsi sel dan baris, seperti `bold:` atau `align:`, juga berlaku untuk teks.
- `tc.p(options) { |p| ... }`: paragraf yang hanya memakai opsinya sendiri.
- `tc.image(media, width:, height:)`: gambar. Lihat masalah ukurannya di bagian [Gambar](#images).
- `tc.page_numbers`: bidang nomor halaman. Masalah yang diketahui: diawali teks "GlobalOptions".

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Gambar

`media = document.add_media(source)` mendaftarkan gambar. `source` bisa berupa:

- objek IO, seperti `File` atau `StringIO`
- path file
- URL `http`, `https`, atau `ftp`

Gambar dibaca saat Anda memanggil `render` dan selalu disimpan sebagai part PNG, jadi file PNG adalah pilihan yang aman.

**Gambar sebaris**: `p.image(media, options)` menempatkan gambar di dalam baris teks.

| Opsi | Nilai | Efek |
|---|---|---|
| `style` | Hash properti CSS | Ukuran, misalnya `{ width: "120pt", height: "60pt" }` |
| `title` | String | Judul gambar |
| `fill_color` | Warna VML seperti `"#00FF00"` | Isian di belakang gambar |
| `opacity` | String seperti `"50%"` | Opasitas isian |
| `wrap` | `true` | Mematikan pembungkusan teks di sekitar gambar |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Gambar berposisi**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` menambatkan gambar dengan offset dari kolom (`pH`) dan dari paragraf (`pV`), dalam EMU. `pH`, `pV`, dan `wrap` harus diberikan semuanya, jika tidak gambar ditempatkan sebaris.

**Gambar header**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` mematikan pembungkusan. Nilai lain mana pun membuat teks membungkus gambar dalam bentuk persegi.

**Masalah yang diketahui di 0.1.0**: gambar berposisi, gambar header, dan gambar di sel tabel menghitung lebar sekaligus tingginya sebagai `width * height * 14.92` EMU. Hasilnya selalu persegi dan biasanya sangat kecil. Gambar sebaris yang ukurannya diatur dengan `style:` tidak terpengaruh. Lihat [Praktik terbaik](Best-Practices#images) untuk solusi sementaranya.

## <a id="headers-and-footers"></a>Header dan footer

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Header dapat berisi `text`, `br`, `table`, `image`, dan `page_numbers`. Footer dapat berisi `text`, `br`, `table`, dan `page_numbers`.
- Opsi yang diberikan ke `Header.new` atau `Footer.new` berlaku untuk teksnya.
- `add_header` dan `add_footer` berlaku mulai halaman 2. Halaman 1 selalu memakai `add_first_header` dan `add_first_footer`. Lihat [Konfigurasi](Configuration#first-page).

## <a id="page-numbers"></a>Nomor halaman

`page_numbers(name: "Page", align: "center")` menulis teks `name` diikuti bidang PAGE. Nomornya ditampilkan dengan Arial 10 pt berwarna abu-abu gelap. Nilai bawaan `align` adalah `"right"`.

Masalah yang diketahui: kode bidang ditulis tanpa garis miring terbalik pada `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Daftar isi

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Jika diberi nomor halaman, entri mendapatkan perhentian tab rata kanan di `pos` (dalam seperdua puluh poin; 9360 sama dengan 6,5 inci) yang diisi karakter `leader`. Nilai karakter pengisi berasal dari WordprocessingML, seperti `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"`, dan `"none"`.
- Tanpa nomor halaman, entri berupa teks setelah satu tab.
- Nomor halamannya adalah yang Anda berikan. docxi tidak membuat bidang TOC Word dan tidak menghitung halaman.

## <a id="hyperlinks"></a>Hyperlink

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` mengembalikan ID relasi, dan `hyperlink(text, url, id)` menempatkan tautannya. Tujuannya diambil dari relasi, jadi berikan ID yang dikembalikan untuk URL yang Anda inginkan. Hanya tautan eksternal yang didukung.

## <a id="frames-text-boxes"></a>Bingkai (kotak teks)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Bingkai adalah kotak teks VML dengan isian abu-abu muda semi-transparan dan garis tepi hitam yang tetap. Teks berada di atas dan di bawahnya.
- Di dalam bingkai, Anda bisa memanggil `text`, `br`, `tab`, `image`, dan `frame`.
- `spacing: true` menambahkan jarak sebelum teks.

## <a id="output"></a>Keluaran

`word.render` mengembalikan `StringIO` yang berisi paket `.docx` lengkap. Arsipnya ditulis tanpa rekaman Zip64, sehingga Google Docs dan versi LibreOffice yang lebih lama bisa membukanya. [Arsitektur](Architecture) mencantumkan setiap part dalam paket.

## <a id="related-pages"></a>Halaman terkait

- [Contoh penggunaan](Usage-Examples)
- [Referensi API](API-Reference)
- [Peta jalan](Roadmap), untuk masalah yang diketahui dan direncanakan untuk 0.2.0

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
