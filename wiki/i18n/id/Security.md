<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Keamanan docxi: melaporkan masalah dan menangani input yang tidak tepercaya

Halaman ini menjelaskan cara melaporkan masalah keamanan, serta hal-hal keamanan yang perlu diketahui saat aplikasi Anda meneruskan input pengguna ke docxi.

## <a id="reporting-a-vulnerability"></a>Melaporkan kerentanan

Mohon jangan melaporkan masalah keamanan di issue publik. Kirim email ke pengelola sebagai gantinya. Alamatnya ada di [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Sertakan versi docxi dan cara singkat untuk mereproduksi masalahnya.

## <a id="upgrade-from-00x"></a>Tingkatkan dari 0.0.x

docxi 0.0.x meneruskan string path gambar ke `Kernel#open`. `Kernel#open` menjalankan string yang diawali `|` sebagai perintah shell, sehingga penyerang yang mengendalikan path gambar bisa menjalankan perintah di server Anda. 0.1.0 membuka file dengan `File.open` dan mengunduh URL dengan `URI#open`. Keduanya tidak menjalankan perintah.

## <a id="untrusted-input"></a>Input yang tidak tepercaya

**Teks aman untuk diberikan.** Teks, isi tabel, dan string lainnya ditulis sebagai teks XML atau nilai atribut, yang di-escape oleh nokogiri. Input pengguna tidak bisa menambahkan elemen ke dokumen.

**Sumber gambar perlu kehati-hatian.** `add_media` menerima:

| Sumber | Yang dilakukan docxi |
|---|---|
| IO yang Anda buka sendiri | Membacanya |
| Path file | Membaca file tersebut dari disk server Anda |
| URL `http`, `https`, atau `ftp` | Mengunduhnya dari server Anda |

Jika pengguna bisa memilih path, mereka bisa membuat file apa pun yang dapat dibaca di server Anda, seperti konfigurasi atau kunci, tertanam di dalam dokumen. Jika pengguna bisa memilih URL, server Anda akan memintanya, dan permintaan itu bisa menjangkau alamat internal (pemalsuan permintaan sisi server, SSRF). Berikan IO yang Anda buka sendiri, atau izinkan hanya path dan host yang dikenal.

**Hyperlink** mengarah ke mana pun URL-nya menunjuk. Periksa URL dari pengguna sebelum memberikannya ke `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Cara rilis dilindungi

- Rilis diterbitkan dari GitHub Actions dengan trusted publishing RubyGems, sehingga tidak ada kunci API yang disimpan di repositori.
- Metadata gem mengatur `rubygems_mfa_required`, sehingga push yang memakai kunci API pribadi memerlukan autentikasi multifaktor.

## <a id="related-pages"></a>Halaman terkait

- [Praktik terbaik](Best-Practices)
- [Deployment](Deployment)
- [Berkontribusi](Contributing)

---

[← Beranda wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
