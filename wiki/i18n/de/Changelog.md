<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Änderungsprotokoll und Versionsverlauf

Diese Seite fasst jede Version zusammen. Die vollständigen Hinweise stehen in [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) im Repository, das bei Abweichungen maßgeblich ist.

## <a id="releases"></a>Versionen

| Version | Datum | Hinweise |
|---|---|---|
| 0.1.1 | 2026-10-29 | Genauere Kurzbeschreibung und Beschreibung des Gems sowie ein Link zur Dokumentation. Keine Codeänderungen |
| 0.1.0 | 2026-10-28 | Funktioniert mit aktuellem Ruby, Rails und rubyzip. Erzeugt dieselben Dokumente wie 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Vorabversion von 0.1.0 |
| 0.0.6 | 2016-10-18 | Letzte Version vor dem Update auf 0.1.0 |
| 0.0.2 | 2016-10-14 | Es wurden keine Versionshinweise festgehalten |
| 0.0.1 | 2016-10-14 | Erste Version. Es wurden keine Versionshinweise festgehalten |

Die Versionen 0.0.3 bis 0.0.5 wurden nie auf RubyGems veröffentlicht.

## <a id="011"></a>0.1.1

Keine Codeänderungen, und die erzeugten Dokumente sind dieselben wie mit 0.1.0.

- Kurzbeschreibung und Beschreibung des Gems sagen jetzt, was docxi tut. Vorher beschrieben sie es als Vorlagengenerator.
- RubyGems verlinkt dieses Wiki als Dokumentation des Gems.

## <a id="010"></a>0.1.0

**Kompatibilität**

- Erfordert Ruby 2.5 oder neuer. Anwendungen mit älterem Ruby bleiben bei 0.0.6.
- Funktioniert mit und ohne Rails und wird zusammen mit Rails 5.2 bis 8.1 getestet.
- nokogiri `>= 1.8, < 2` statt `>= 1.6.3, < 1.7`, was mit jeder Rails-Version seit 6.0 kollidierte.
- rubyzip `>= 1.3, < 4` statt `~> 0.9.9`.
- Ruft nicht mehr `blank?` und `present?` von ActiveSupport auf, die nur funktionierten, wenn Rails sie geladen hatte.

**Korrekturen**

- Mit rubyzip 3 geschriebene Dokumente enthalten keine Zip64-Einträge mehr, die Google Docs und ältere LibreOffice-Versionen ablehnen.
- `add_media` mit einer Bild-URL funktioniert unter Ruby 3 wieder.
- `add_media` führt einen String, der mit `|` beginnt, nicht mehr als Shell-Befehl aus.
- `tab` und `frame` lösen keinen `NameError` mehr aus.
- `document.frame("text")` erstellt ein Textfeld mit dem Text.
- `page_numbers` in einer Kopfzeile löst keinen `NameError` mehr aus.

**Gem**

- Neue Projektpflege sowie neue Links zur Homepage und zum Änderungsprotokoll.
- Veröffentlichungen erfordern MFA.

## <a id="related-pages"></a>Verwandte Seiten

- [Roadmap](Roadmap)
- [Installation](Installation#upgrading-from-00x)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
