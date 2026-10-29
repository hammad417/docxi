<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# An docxi mitwirken

Fehlerberichte, Korrekturen, Dokumentation und Übersetzungen sind willkommen. Diese Seite erklärt, wie Sie ein Problem melden und was ein Pull Request braucht. Wie Sie eine Arbeitskopie einrichten, steht unter [Entwicklungsumgebung](Development-Setup).

## <a id="reporting-a-bug"></a>Einen Fehler melden

Eröffnen Sie ein Issue unter [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) mit:

- einem kurzen Ruby-Skript, das das Problem reproduziert
- den Versionen von Ruby, docxi, nokogiri und rubyzip (und Rails, falls Sie es verwenden)
- dem, was Sie erwartet haben, und dem, was Sie erhalten haben. Bei fehlerhaften Dokumenten nennen Sie das Programm, mit dem Sie sie geöffnet haben

Sicherheitsprobleme melden Sie bitte vertraulich. Siehe [Sicherheit](Security).

## <a id="pull-requests"></a>Pull Requests

1. Forken Sie das Repository und legen Sie einen Branch an.
2. Nehmen Sie die Änderung vor und halten Sie sich an den Stil des umgebenden Codes.
3. Fügen Sie Tests hinzu oder aktualisieren Sie sie und führen Sie `bundle exec rake test` aus.
4. Wenn sich das erzeugte XML absichtlich ändert, führen Sie `bundle exec rake golden` aus und erklären Sie die Änderung im Pull Request. Die Referenzdateien schützen die Ausgabe, auf die sich bestehende Nutzer verlassen.
5. Aktualisieren Sie `CHANGELOG.md` und die Dokumentation, wenn sich das Verhalten ändert.
6. Eröffnen Sie den Pull Request. Die CI führt die Tests unter Ruby 2.5 bis 4.0 und Rails 5.2 bis 8.1 aus und prüft, ob LibreOffice die Beispieldokumente öffnet.

Änderungen, die bestehende Dokumente verändern, etwa die für 0.2.0 geplanten Korrekturen, gehören in einen eigenen Pull Request, damit sich der Diff der Referenzdateien leicht prüfen lässt.

## <a id="documentation-and-translations"></a>Dokumentation und Übersetzungen

- Die englische README ist die Quelle von 15 Übersetzungen in `docs/i18n/`. Folgen Sie bei Änderungen `docs/i18n/TRANSLATION_GUIDE.md` im Repository.
- Dieses Wiki wird im Ordner `wiki/` des Repositorys geschrieben und von dort veröffentlicht. `wiki/README.md` erklärt, wie.

## <a id="related-pages"></a>Verwandte Seiten

- [Entwicklungsumgebung](Development-Setup)
- [Architektur](Architecture)
- [Roadmap](Roadmap)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
