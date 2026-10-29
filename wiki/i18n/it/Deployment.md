<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# docxi in produzione e pubblicazione delle versioni

docxi è una libreria, non un servizio, quindi non c'è nulla da distribuire a sé. Questa pagina spiega cosa conta quando rilasci un'applicazione che usa docxi e come vengono pubblicate le nuove versioni di docxi.

## <a id="shipping-an-app-that-uses-docxi"></a>Rilasciare un'applicazione che usa docxi

- **Dipendenze.** `bundle install` installa nokogiri e rubyzip. nokogiri fornisce gem precompilate per le piattaforme Linux, macOS e Windows più diffuse. Su altre piattaforme, per esempio alcune immagini di container minimali, viene compilato dai sorgenti e serve un compilatore C. Consulta la [guida all'installazione di nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Nessun software da ufficio.** Server e container non hanno bisogno di Microsoft Word né di LibreOffice.
- **Memoria.** `render` costruisce l'intero documento in memoria e lo restituisce come `StringIO`. L'uso della memoria cresce con il numero di pagine e con le dimensioni delle immagini.
- **Rete e file.** docxi legge il file system solo per i percorsi delle immagini che passi ad `add_media`, e usa la rete solo per gli URL delle immagini che passi. Consulta [Sicurezza](Security).
- **Documenti lunghi.** Generali in un job in background anziché durante una richiesta web. Consulta [Integrazioni](Integrations).

## <a id="how-docxi-releases-are-published"></a>Come vengono pubblicate le versioni di docxi

Questi passaggi riguardano chi cura il progetto. Le versioni arrivano su RubyGems tramite GitHub Actions con la pubblicazione attendibile (trusted publishing), quindi nel repository non è conservata alcuna chiave API.

1. Aggiorna `lib/docxi/version.rb` e `CHANGELOG.md`, poi esegui il merge in `master`.
2. Applica al commit un tag con la versione preceduta da `v`, per esempio `v0.1.1`, ed esegui il push del tag.
3. `.github/workflows/release.yml` compila la gem e la pubblica su RubyGems.

Le versioni preliminari seguono gli stessi passaggi con una versione come `0.2.0.rc1`. Bundler installa le versioni preliminari solo se vengono richieste esplicitamente. La CI (`.github/workflows/ci.yml`) viene eseguita a ogni push su `master` e per ogni pull request.

## <a id="related-pages"></a>Pagine correlate

- [Installazione](Installation)
- [Ambiente di sviluppo](Development-Setup)
- [Registro delle modifiche](Changelog)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
