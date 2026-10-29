<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Registro delle modifiche e storia delle versioni di docxi

Questa pagina riassume ogni versione. Le note complete si trovano in [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) nel repository, che fa fede se le due fonti dovessero differire.

## <a id="releases"></a>Versioni

| Versione | Data | Note |
|---|---|---|
| 0.1.1 | 2026-10-29 | Sommario e descrizione della gem più precisi, e link alla documentazione. Nessuna modifica al codice |
| 0.1.0 | 2026-10-28 | Funziona con le versioni attuali di Ruby, Rails e rubyzip. Genera gli stessi documenti della 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Versione preliminare della 0.1.0 |
| 0.0.6 | 2016-10-18 | Ultima versione prima dell'aggiornamento 0.1.0 |
| 0.0.2 | 2016-10-14 | Non sono state registrate note di versione |
| 0.0.1 | 2016-10-14 | Prima versione. Non sono state registrate note di versione |

Le versioni dalla 0.0.3 alla 0.0.5 non sono mai state pubblicate su RubyGems.

## <a id="011"></a>0.1.1

Nessuna modifica al codice, e i documenti generati sono identici a quelli della 0.1.0.

- Il sommario e la descrizione della gem ora dicono cosa fa docxi. Prima lo presentavano come un generatore di modelli.
- RubyGems collega questa wiki come documentazione della gem.

## <a id="010"></a>0.1.0

**Compatibilità**

- Richiede Ruby 2.5 o successivo. Le applicazioni con versioni di Ruby più vecchie restano sulla 0.0.6.
- Funziona con o senza Rails, ed è testata insieme a Rails dalla 5.2 alla 8.1.
- nokogiri `>= 1.8, < 2` invece di `>= 1.6.3, < 1.7`, che andava in conflitto con ogni versione di Rails dalla 6.0 in poi.
- rubyzip `>= 1.3, < 4` invece di `~> 0.9.9`.
- Non chiama più `blank?` e `present?` di ActiveSupport, che funzionavano solo quando Rails li aveva caricati.

**Correzioni**

- I documenti scritti con rubyzip 3 non contengono più record Zip64, che Google Docs e le versioni meno recenti di LibreOffice rifiutano.
- `add_media` con l'URL di un'immagine funziona di nuovo con Ruby 3.
- `add_media` non esegue più come comando di shell una stringa che inizia con `|`.
- `tab` e `frame` non sollevano più `NameError`.
- `document.frame("text")` crea una casella di testo che contiene il testo.
- `page_numbers` in un'intestazione non solleva più `NameError`.

**Gem**

- Nuova gestione della manutenzione e nuovi link alla homepage e al registro delle modifiche.
- Le pubblicazioni richiedono la MFA.

## <a id="related-pages"></a>Pagine correlate

- [Roadmap](Roadmap)
- [Installazione](Installation#upgrading-from-00x)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
