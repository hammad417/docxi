<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# docxi in Produktion betreiben und Versionen veröffentlichen

docxi ist eine Bibliothek, kein Dienst, daher gibt es für sich genommen nichts zu deployen. Diese Seite beschreibt, worauf es ankommt, wenn Sie eine Anwendung mit docxi ausliefern, und wie neue docxi-Versionen veröffentlicht werden.

## <a id="shipping-an-app-that-uses-docxi"></a>Eine Anwendung mit docxi ausliefern

- **Abhängigkeiten.** `bundle install` installiert nokogiri und rubyzip. nokogiri liefert vorkompilierte Gems für gängige Linux-, macOS- und Windows-Plattformen. Auf anderen Plattformen, etwa manchen minimalen Container-Images, wird es aus dem Quellcode kompiliert und braucht einen C-Compiler. Siehe die [Installationsanleitung von nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Keine Office-Software.** Server und Container brauchen weder Microsoft Word noch LibreOffice.
- **Arbeitsspeicher.** `render` baut das ganze Dokument im Speicher auf und gibt es als `StringIO` zurück. Der Speicherbedarf wächst mit der Seitenzahl und der Größe der Bilder.
- **Netzwerk und Dateien.** docxi liest das Dateisystem nur für Bildpfade, die Sie an `add_media` übergeben, und nutzt das Netzwerk nur für Bild-URLs, die Sie übergeben. Siehe [Sicherheit](Security).
- **Lange Dokumente.** Erzeugen Sie sie in einem Hintergrundjob statt während eines Web-Requests. Siehe [Integrationen](Integrations).

## <a id="how-docxi-releases-are-published"></a>Wie docxi-Versionen veröffentlicht werden

Diese Schritte richten sich an die Projektpflege. Versionen gelangen über GitHub Actions mit Trusted Publishing zu RubyGems, daher ist im Repository kein API-Schlüssel gespeichert.

1. Aktualisieren Sie `lib/docxi/version.rb` und `CHANGELOG.md` und mergen Sie nach `master`.
2. Taggen Sie den Commit mit der Version mit vorangestelltem `v`, zum Beispiel `v0.1.1`, und pushen Sie den Tag.
3. `.github/workflows/release.yml` baut das Gem und veröffentlicht es auf RubyGems.

Vorabversionen folgen denselben Schritten mit einer Version wie `0.2.0.rc1`. Bundler installiert Vorabversionen nur, wenn sie ausdrücklich angefordert werden. Die CI (`.github/workflows/ci.yml`) läuft bei jedem Push auf `master` und bei jedem Pull Request.

## <a id="related-pages"></a>Verwandte Seiten

- [Installation](Installation)
- [Entwicklungsumgebung](Development-Setup)
- [Änderungsprotokoll](Changelog)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
