<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# docxi installieren: Voraussetzungen, Bundler-Einrichtung und Upgrade

docxi ist ein reines Ruby-Gem. Bei der Installation kommen nur zwei Abhängigkeiten hinzu, nokogiri und rubyzip, sonst nichts. Eine Office-Software ist nicht nötig.

## <a id="requirements"></a>Voraussetzungen

| Voraussetzung | Version |
|---|---|
| Ruby | 2.5 oder neuer |
| nokogiri | `>= 1.8, < 2`, wird automatisch installiert |
| rubyzip | `>= 1.3, < 4`, wird automatisch installiert |
| Rails | Nicht erforderlich |

## <a id="install-with-bundler"></a>Mit Bundler installieren

Fügen Sie das Gem zu Ihrem Gemfile hinzu:

```ruby
gem "docxi", "~> 0.1"
```

Installieren Sie es anschließend:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Ohne Bundler installieren

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Installation prüfen

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

In einem Bundler-Projekt führen Sie stattdessen `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` aus. Der Befehl gibt die installierte Version aus, zum Beispiel `0.1.0`.

## <a id="upgrading-from-00x"></a>Upgrade von 0.0.x

Wenn Ihr Gemfile docxi auf eine 0.0.x-Version festlegt, zum Beispiel mit `"~> 0.0.6"`, ändern Sie die Zeile in `gem "docxi", "~> 0.1"`. Führen Sie dann `bundle update docxi` aus. Codeänderungen sind nicht nötig, und 0.1.0 erzeugt dieselben Dokumente wie 0.0.6: Die Testsuite vergleicht die Ausgabe mit Dateien, die 0.0.6 erzeugt hat.

Was sich intern geändert hat:

- nokogiri und rubyzip dürfen jetzt jede Version aus den oben genannten Bereichen haben, sodass sich docxi neben aktuellen Rails-Versionen installieren lässt.
- docxi setzt kein geladenes ActiveSupport mehr voraus.
- Archive werden ohne Zip64-Einträge geschrieben.
- Bilder werden ohne `Kernel#open` geladen, daher funktionieren URLs unter Ruby 3.
- `tab`, `frame` und Seitenzahlen in Kopfzeilen funktionieren.

Anwendungen mit einer Ruby-Version unter 2.5 erhalten weiterhin 0.0.6, weil 0.1.0 Ruby 2.5 als Mindestversion angibt. Das [Änderungsprotokoll](Changelog) listet alle Änderungen auf.

## <a id="versions-tested-in-ci"></a>In der CI getestete Versionen

Jede Rails-Version läuft mit vollständig geladenem Framework neben docxi.

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

Die CI führt außerdem aus:

- die Tests ohne Rails unter Ruby 2.5, 2.6 und 4.0
- die Tests mit rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) und nokogiri 1.11 (Ruby 2.5)
- eine Prüfung, ob LibreOffice die Beispieldokumente konvertieren kann

## <a id="platform-notes"></a>Hinweise zu Plattformen

- nokogiri liefert vorkompilierte Gems für gängige Linux-, macOS- und Windows-Plattformen. Passt keines davon, wird es aus dem Quellcode kompiliert und braucht einen C-Compiler. Siehe die [Installationsanleitung von nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- Die CI läuft unter Ubuntu 24.04, und die Tests wurden auch unter macOS ausgeführt. Windows wird nicht getestet.

## <a id="if-installation-fails"></a>Wenn die Installation fehlschlägt

Bundler-Konflikte bei nokogiri oder rubyzip bedeuten fast immer, dass ein altes docxi 0.0.6 aufgelöst wird. Siehe [Fehlerbehebung](Troubleshooting#installation).

## <a id="related-pages"></a>Verwandte Seiten

- [Erste Schritte](Getting-Started)
- [Konfiguration](Configuration)
- [Deployment](Deployment)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
