<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Entwicklungsumgebung: Tests und Werkzeuge von docxi ausführen

Diese Seite bringt eine Arbeitskopie von docxi zum Laufen und erklärt die Tests, die Beispieldokumente und die CI-Einrichtung.

## <a id="get-the-code"></a>Den Code holen

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Für die Entwicklung eignet sich jedes Ruby ab 2.5. Die Entwicklungsabhängigkeiten sind minitest und rake.

## <a id="rake-tasks"></a>Rake-Tasks

| Befehl | Was er tut |
|---|---|
| `bundle exec rake test` | Führt die Testsuite aus (der Standard-Task) |
| `bundle exec rake samples` | Schreibt die Beispieldokumente nach `tmp/samples/*.docx`, zum Öffnen in Word, LibreOffice oder Google Docs |
| `bundle exec rake golden` | Schreibt `test/golden` aus dem aktuellen Code neu. Nur nach einer beabsichtigten Änderung der Ausgabe verwenden |
| `bundle exec rake build` | Baut das Gem nach `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Die Tests mit geladenem Rails ausführen

Der Ordner `gemfiles/` enthält ein Gemfile pro Rails-Version sowie einige für ältere rubyzip- und nokogiri-Versionen:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` lädt vor den Tests das ganze Framework. Außerdem vergleicht es die `blank?`-Regeln von docxi mit denen von ActiveSupport.

## <a id="what-the-tests-cover"></a>Was die Tests abdecken

| Datei | Prüft |
|---|---|
| `test/test_golden_output.rb` | Jeder Teil zweier Beispieldokumente entspricht dem, was 0.0.6 erzeugt hat |
| `test/test_package.rb` | Das XML ist wohlgeformt, Beziehungen zeigen auf vorhandene Teile, jeder Teil hat einen Inhaltstyp und es gibt keine Zip64-Einträge |
| `test/test_media.rb` | Bilder aus Pfaden, IO-Objekten und URLs, und dass ein `\|`-String nie als Befehl ausgeführt wird |
| `test/test_fixed_features.rb` | `tab`, `frame` und Seitenzahlen in Kopfzeilen |
| `test/test_util.rb` | `Docxi::Util.blank?` folgt den Regeln von ActiveSupport |
| `test/test_rails_independence.rb` | docxi läuft ohne ActiveSupport |
| `test/test_warnings.rb` | Das Rendern gibt keine Deprecation-Warnungen aus `lib/` aus |

Die Beispieldokumente liegen in `test/support/sample_documents.rb`, und `test/support/package.rb` liest eine `.docx`-Datei direkt aus ihren Zip-Strukturen.

## <a id="continuous-integration"></a>Continuous Integration

`.github/workflows/ci.yml` läuft bei jedem Push auf `master` und bei jedem Pull Request. Abgedeckt sind:

- Ruby 2.5, 2.6 und 4.0 ohne Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) und 8.1 (4.0)
- Rails `main`, das fehlschlagen darf
- rubyzip 1.3 und 2.4 sowie nokogiri 1.11
- ein Job, der die Beispieldokumente mit LibreOffice in PDF konvertiert

Veröffentlichungen werden unter [Deployment](Deployment) beschrieben.

## <a id="related-pages"></a>Verwandte Seiten

- [Mitwirken](Contributing)
- [Architektur](Architecture)
- [Deployment](Deployment)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
