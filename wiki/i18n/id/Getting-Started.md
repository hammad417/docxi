<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Memulai docxi: dokumen Word pertama Anda dengan Ruby

Panduan ini membawa Anda dari folder kosong sampai ke file `.docx` yang berformat. Panduan ini mengasumsikan Anda memahami dasar-dasar Ruby. Anda tidak memerlukan Rails atau Microsoft Word.

## <a id="1-install-the-gem"></a>1. Pasang gem

```sh
gem install docxi
```

Di proyek yang memakai Gemfile, tambahkan `gem "docxi", "~> 0.1"` lalu jalankan `bundle install` sebagai gantinya. Detailnya ada di [Instalasi](Installation).

## <a id="2-create-a-document"></a>2. Buat dokumen

Simpan kode berikut sebagai `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Jalankan `ruby hello.rb` (atau `bundle exec ruby hello.rb` di proyek Bundler), lalu buka `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Fungsi setiap baris

- `Docxi::Document.new` membuat dokumen kosong dengan margin halaman bawaan.
- `word.document` adalah isi dokumen. Semua yang Anda tambahkan melewatinya.
- `document.p` menambahkan paragraf. Opsi yang diberikan ke `p`, seperti `align:`, berlaku untuk paragraf dan teks di dalamnya.
- `p.text` menambahkan satu run teks. `size:` dinyatakan dalam poin.
- `word.render` menyusun paket `.docx` dan mengembalikan sebuah `StringIO`. `.string` miliknya berisi byte file tersebut.

## <a id="4-add-a-list-and-a-table"></a>4. Tambahkan daftar dan tabel

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

Ada dua hal yang perlu diperhatikan:

- `space: true` mempertahankan spasi di awal atau akhir sebuah run. Tanpa opsi ini, Word membuang spasi tersebut dan kata-kata jadi menyambung.
- Opsi pada `tr` berlaku untuk setiap sel di baris itu, sehingga seluruh baris judul tabel diberi arsiran latar dan dicetak tebal.

## <a id="5-where-to-go-next"></a>5. Langkah selanjutnya

- [Contoh penggunaan](Usage-Examples): resep lengkap untuk laporan, faktur, dan surat
- [Fitur](Features): semua opsi untuk paragraf, teks, tabel, gambar, header, dan footer
- [Integrasi](Integrations): mengirim dokumen dari controller Rails
- [Praktik terbaik](Best-Practices): menghindari kejutan umum sebelum terjadi

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
