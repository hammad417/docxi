<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Peta jalan docxi dan masalah yang diketahui

docxi tidak punya jadwal rilis maupun peta jalan yang sudah dijanjikan. Satu-satunya pekerjaan terencana yang tercatat adalah memperbaiki masalah yang diketahui di bawah ini. Perbaikannya direncanakan untuk 0.2.0, tanpa tanggal. Perbaikan ini tidak dimasukkan ke 0.1.0 karena akan mengubah dokumen yang sudah dihasilkan oleh pengguna saat ini.

## <a id="known-issues-planned-for-020"></a>Masalah yang diketahui dan direncanakan untuk 0.2.0

| Masalah | Solusi sementara di 0.1.0 |
|---|---|
| Gambar berposisi, gambar header, dan gambar sel tabel menghitung ukurannya sebagai `width * height * 14.92` EMU untuk kedua dimensi | Gunakan gambar sebaris yang ukurannya diatur dengan `style:`, atau lihat [Praktik terbaik](Best-Practices#images) |
| `borders:` tingkat tabel diterapkan ke setiap sel, bukan ke tepi luar tabel | Untuk garis luar saja, atur `borders:` pada sisi luar sel-sel di tepi tabel |
| Kode bidang nomor halaman ditulis tanpa garis miring terbalik pada `\* MERGEFORMAT` | Tidak ada |
| Nomor halaman di sel tabel diawali "GlobalOptions" | Taruh nomor halaman di header atau footer |
| Gambar selalu disimpan sebagai part PNG, dan `[Content_Types].xml` hanya mendeklarasikan dua header dan dua footer pertama | Gunakan gambar PNG, dan paling banyak dua header dan dua footer |
| Memberikan hanya sebagian margin membuat sisanya ditulis kosong | Berikan `top`, `right`, `bottom`, dan `left` sekaligus |
| `docProps/core.xml` mencatat waktu lokal dengan penanda UTC dan tidak pernah menulis nama penulis | Tidak ada |
| `Docxi.test` memunculkan `NoMethodError` | Jangan memanggilnya. Itu helper sisa dari versi lama |
| `document.br` dan `document.tab` tanpa `page: true` menghasilkan XML yang tidak valid | Gunakan `p.br` dan `p.tab` di dalam paragraf |

## <a id="not-planned"></a>Tidak direncanakan

Butir-butir ini tidak ada dalam daftar di atas, dan belum ada keputusan apa pun mengenainya:

- mengedit dokumen yang sudah ada atau mengisi templat
- daftar bernomor, gaya judul, atau daftar isi otomatis
- ukuran halaman lain atau halaman lanskap

## <a id="related-pages"></a>Halaman terkait

- [Catatan perubahan](Changelog)
- [Berkontribusi](Contributing)
- [Tanya jawab](FAQ)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
