<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Referensi API docxi: kelas dan metode

docxi adalah pustaka Ruby, jadi halaman ini mencantumkan API Ruby-nya. Tidak ada API HTTP. Opsi setiap metode dijelaskan di [Fitur](Features), dan kelas-kelasnya dijelaskan di [Arsitektur](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Metode | Mengembalikan | Catatan |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` adalah margin halaman. Lihat [Konfigurasi](Configuration) |
| `#document` | `Docxi::Word::Document` | Isi dokumen tempat Anda menambahkan konten |
| `#render` | `StringIO` | Paket `.docx` lengkap. `.string` memberikan byte-nya |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Metode | Mengembalikan | Menambahkan |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Paragraf |
| `text(string, options = {})` | Paragraph | Paragraf dengan satu run teks |
| `br(page: true)` | Break | Pemisah halaman |
| `tab(page: true)` | Tab | Paragraf berisi satu tab |
| `table(options = {}) { \|table\| }` | Table | Tabel |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Entri daftar isi |
| `frame(text, options = {})` | Paragraph | Paragraf berisi kotak teks |
| `add_media(io_path_or_url, options = {})` | Media | Gambar untuk ditempatkan dengan `image` |
| `add_hyperlink(url, options = {})` | String | ID relasi untuk `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Header untuk halaman 2 dan seterusnya, serta untuk halaman 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Footer untuk halaman 2 dan seterusnya, serta untuk halaman 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Paragraf (`Docxi::Word::Contents::Paragraph`)

| Metode | Menambahkan |
|---|---|
| `text(string, options = {})` | Satu run teks. Opsi paragraf menjadi nilai bawaannya |
| `br(times: 1)` | Pemisah baris |
| `tab(times: 1)` | Karakter tab |
| `image(media, options = {})` | Gambar sebaris atau berposisi |
| `hyperlink(text, url, id)` | Tautan yang memakai `id` relasi dari `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Kotak teks |

## <a id="tables"></a>Tabel

| Metode | Menambahkan |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Baris. Opsi tabel menjadi nilai bawaannya |
| `TableRow#tc(options = {}) { \|tc\| }` | Sel. Opsi baris menjadi nilai bawaannya |
| `TableCell#text(string, options = {})` | Paragraf. Opsi sel menjadi nilai bawaannya |
| `TableCell#p(options = {}) { \|p\| }` | Paragraf dengan opsinya sendiri |
| `TableCell#image(media, width:, height:)` | Gambar |
| `TableCell#page_numbers` | Bidang nomor halaman |

## <a id="table-of-contents"></a>Daftar isi

| Metode | Menambahkan |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Satu entri daftar isi |

## <a id="headers-and-footers"></a>Header dan footer

| Kelas | Metode |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Bingkai

`Docxi::Word::Contents::Frame` menerima `text`, `br`, `tab`, `image`, dan `frame` di dalam bloknya.

## <a id="constants-and-helpers"></a>Konstanta dan helper

| Nama | Arti |
|---|---|
| `Docxi::VERSION` | Versi yang terpasang, misalnya `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Menulis satu part ke arsip dengan ukurannya diatur di awal. Dipakai secara internal oleh setiap part |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Helper internal dengan aturan `blank?` milik ActiveSupport |

## <a id="related-pages"></a>Halaman terkait

- [Fitur](Features)
- [Contoh penggunaan](Usage-Examples)
- [Arsitektur](Architecture)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
