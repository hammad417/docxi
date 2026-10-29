<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Ambiente di sviluppo: eseguire i test e gli strumenti di docxi

Questa pagina ti aiuta a far funzionare una copia di lavoro di docxi e spiega i test, i documenti di esempio e la configurazione della CI.

## <a id="get-the-code"></a>Ottenere il codice

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Per lo sviluppo va bene qualsiasi Ruby dalla 2.5 in su. Le dipendenze di sviluppo sono minitest e rake.

## <a id="rake-tasks"></a>Task di Rake

| Comando | Cosa fa |
|---|---|
| `bundle exec rake test` | Esegue la suite di test (il task predefinito) |
| `bundle exec rake samples` | Scrive i documenti di esempio in `tmp/samples/*.docx`, da aprire in Word, LibreOffice o Google Docs |
| `bundle exec rake golden` | Riscrive `test/golden` a partire dal codice attuale. Usalo solo dopo una modifica intenzionale dell'output |
| `bundle exec rake build` | Compila la gem in `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Eseguire i test con Rails caricato

La cartella `gemfiles/` contiene un Gemfile per ogni versione di Rails, più alcuni per le versioni meno recenti di rubyzip e nokogiri:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` carica l'intero framework prima dei test. Confronta anche le regole di `blank?` di docxi con quelle di ActiveSupport.

## <a id="what-the-tests-cover"></a>Cosa coprono i test

| File | Verifica |
|---|---|
| `test/test_golden_output.rb` | Ogni parte di due documenti di esempio corrisponde a ciò che generava la 0.0.6 |
| `test/test_package.rb` | L'XML è ben formato, le relazioni puntano a parti esistenti, ogni parte ha un tipo di contenuto e non ci sono record Zip64 |
| `test/test_media.rb` | Le immagini da percorsi, oggetti IO e URL, e il fatto che una stringa con `\|` non venga mai eseguita come comando |
| `test/test_fixed_features.rb` | `tab`, `frame` e i numeri di pagina nelle intestazioni |
| `test/test_util.rb` | `Docxi::Util.blank?` segue le regole di ActiveSupport |
| `test/test_rails_independence.rb` | docxi funziona senza ActiveSupport |
| `test/test_warnings.rb` | La generazione non emette avvisi di deprecazione da `lib/` |

I documenti di esempio si trovano in `test/support/sample_documents.rb`, e `test/support/package.rb` legge un `.docx` direttamente dalle sue strutture zip.

## <a id="continuous-integration"></a>Integrazione continua

`.github/workflows/ci.yml` viene eseguito a ogni push su `master` e per ogni pull request. Copre:

- Ruby 2.5, 2.6 e 4.0 senza Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) e 8.1 (4.0)
- Rails `main`, a cui è consentito fallire
- rubyzip 1.3 e 2.4, e nokogiri 1.11
- un job che converte i documenti di esempio in PDF con LibreOffice

Le versioni sono trattate in [Deployment](Deployment).

## <a id="related-pages"></a>Pagine correlate

- [Contribuire](Contributing)
- [Architettura](Architecture)
- [Deployment](Deployment)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
