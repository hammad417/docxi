<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Installare docxi: requisiti, configurazione di Bundler e aggiornamento

docxi è una gem scritta interamente in Ruby. Installandola si aggiungono solo due dipendenze, nokogiri e rubyzip, e nient'altro. Non serve alcun software da ufficio.

## <a id="requirements"></a>Requisiti

| Requisito | Versione |
|---|---|
| Ruby | 2.5 o successivo |
| nokogiri | `>= 1.8, < 2`, installato automaticamente |
| rubyzip | `>= 1.3, < 4`, installato automaticamente |
| Rails | Non richiesto |

## <a id="install-with-bundler"></a>Installazione con Bundler

Aggiungi la gem al tuo Gemfile:

```ruby
gem "docxi", "~> 0.1"
```

Poi installala:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Installazione senza Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Verificare l'installazione

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

In un progetto Bundler, eseguilo come `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. Il comando stampa la versione installata, per esempio `0.1.0`.

## <a id="upgrading-from-00x"></a>Aggiornare dalla 0.0.x

Se il tuo Gemfile fissa docxi a una 0.0.x, per esempio con `"~> 0.0.6"`, cambia la riga in `gem "docxi", "~> 0.1"`. Poi esegui `bundle update docxi`. Non serve modificare il codice, e la 0.1.0 genera gli stessi documenti della 0.0.6: la suite di test confronta l'output con file generati dalla 0.0.6.

Cosa è cambiato sotto il cofano:

- nokogiri e rubyzip possono ora avere qualsiasi versione negli intervalli indicati sopra, quindi docxi si installa accanto alle versioni attuali di Rails.
- docxi non richiede più che ActiveSupport sia caricato.
- Gli archivi vengono scritti senza record Zip64.
- Le immagini vengono caricate senza `Kernel#open`, quindi gli URL funzionano con Ruby 3.
- `tab`, `frame` e i numeri di pagina nelle intestazioni funzionano.

Le applicazioni con un Ruby precedente alla 2.5 continuano a ricevere la 0.0.6, perché la 0.1.0 dichiara Ruby 2.5 come versione minima. Il [registro delle modifiche](Changelog) elenca ogni modifica.

## <a id="versions-tested-in-ci"></a>Versioni testate in CI

Ogni versione di Rails viene eseguita con il framework completo caricato insieme a docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

La CI esegue anche:

- i test senza Rails su Ruby 2.5, 2.6 e 4.0
- i test con rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) e nokogiri 1.11 (Ruby 2.5)
- una verifica che LibreOffice riesca a convertire i documenti di esempio

## <a id="platform-notes"></a>Note sulle piattaforme

- nokogiri fornisce gem precompilate per le piattaforme Linux, macOS e Windows più diffuse. Se nessuna corrisponde, viene compilato dai sorgenti e serve un compilatore C. Consulta la [guida all'installazione di nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- La CI gira su Ubuntu 24.04, e i test sono stati eseguiti anche su macOS. Windows non viene testato.

## <a id="if-installation-fails"></a>Se l'installazione non riesce

I conflitti di Bundler su nokogiri o rubyzip significano quasi sempre che si sta risolvendo una vecchia docxi 0.0.6. Consulta [Risoluzione dei problemi](Troubleshooting#installation).

## <a id="related-pages"></a>Pagine correlate

- [Primi passi](Getting-Started)
- [Configurazione](Configuration)
- [Deployment](Deployment)

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
