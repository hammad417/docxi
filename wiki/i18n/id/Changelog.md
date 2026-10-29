<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Catatan perubahan dan riwayat rilis docxi

Halaman ini merangkum setiap rilis. Catatan lengkapnya ada di [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) di repositori, yang menjadi acuan jika keduanya berbeda.

## <a id="releases"></a>Rilis

| Versi | Tanggal | Catatan |
|---|---|---|
| 0.1.1 | 2026-10-29 | Ringkasan dan deskripsi gem yang lebih tepat, serta tautan ke dokumentasi. Tidak ada perubahan kode |
| 0.1.0 | 2026-10-28 | Berfungsi dengan Ruby, Rails, dan rubyzip terkini. Menghasilkan dokumen yang sama dengan 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Pra-rilis 0.1.0 |
| 0.0.6 | 2016-10-18 | Rilis terakhir sebelum pembaruan 0.1.0 |
| 0.0.2 | 2016-10-14 | Tidak ada catatan rilis yang tercatat |
| 0.0.1 | 2016-10-14 | Rilis pertama. Tidak ada catatan rilis yang tercatat |

Versi 0.0.3 sampai 0.0.5 tidak pernah diterbitkan ke RubyGems.

## <a id="011"></a>0.1.1

Tidak ada perubahan kode, dan dokumen yang dihasilkan sama dengan 0.1.0.

- Ringkasan dan deskripsi gem kini menjelaskan apa yang dilakukan docxi. Sebelumnya keduanya menyebutnya pembuat templat.
- RubyGems menautkan wiki ini sebagai dokumentasi gem.

## <a id="010"></a>0.1.0

**Kompatibilitas**

- Memerlukan Ruby 2.5 atau lebih baru. Aplikasi dengan Ruby yang lebih lama tetap memakai 0.0.6.
- Berfungsi dengan atau tanpa Rails, dan diuji bersama Rails 5.2 sampai 8.1.
- nokogiri `>= 1.8, < 2`, menggantikan `>= 1.6.3, < 1.7` yang berbenturan dengan setiap versi Rails sejak 6.0.
- rubyzip `>= 1.3, < 4`, menggantikan `~> 0.9.9`.
- Tidak lagi memanggil `blank?` dan `present?` milik ActiveSupport, yang hanya berfungsi jika Rails sudah memuatnya.

**Perbaikan**

- Dokumen yang ditulis dengan rubyzip 3 tidak lagi berisi rekaman Zip64, yang ditolak oleh Google Docs dan versi LibreOffice yang lebih lama.
- `add_media` dengan URL gambar kembali berfungsi di Ruby 3.
- `add_media` tidak lagi menjalankan string yang diawali `|` sebagai perintah shell.
- `tab` dan `frame` tidak lagi memunculkan `NameError`.
- `document.frame("text")` membuat kotak teks yang berisi teks tersebut.
- `page_numbers` di header tidak lagi memunculkan `NameError`.

**Gem**

- Pengelola, homepage, dan tautan catatan perubahan yang baru.
- Rilis memerlukan MFA.

## <a id="related-pages"></a>Halaman terkait

- [Peta jalan](Roadmap)
- [Instalasi](Installation#upgrading-from-00x)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
