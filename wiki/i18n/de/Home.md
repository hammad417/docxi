<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Wiki: Word-Dokumente (.docx) mit Ruby erzeugen

docxi ist ein Ruby-Gem, das Microsoft-Word-Dateien im Format `.docx` aus Code erzeugt. Es unterstützt Absätze, formatierten Text, Tabellen, Bilder, Kopf- und Fußzeilen, Seitenzahlen, Hyperlinks und ein einfaches Inhaltsverzeichnis. Es braucht weder Microsoft Word noch LibreOffice noch eine Vorlage und funktioniert in reinem Ruby ebenso wie in Ruby-on-Rails-Anwendungen.

Die [README](https://github.com/hammad417/docxi#readme) ist die kurze Einführung. Dieses Wiki beschreibt alle Optionen, enthält Rezepte zum Kopieren, erklärt, wie docxi intern arbeitet, und hilft, wenn etwas nicht funktioniert.

Aktuelle Version: **0.1.1**, für Ruby 2.5 oder neuer. Siehe das [Änderungsprotokoll](Changelog).

## <a id="start-here"></a>Hier anfangen

**Neu bei docxi**

- [Erste Schritte](Getting-Started): das Gem installieren und das erste Dokument erstellen
- [Installation](Installation): Voraussetzungen, Bundler und Upgrade von 0.0.x
- [Anwendungsbeispiele](Usage-Examples): Berichte, Rechnungen, Briefe und Tabellen aus Daten, fertig zum Kopieren

**Dokumente erstellen**

- [Funktionen](Features): jedes Element und die Optionen, die es akzeptiert
- [Konfiguration](Configuration): Seitenränder, Seiteneinrichtung und Einheiten
- [API-Referenz](API-Reference): Klassen und Methoden im Überblick
- [Best Practices](Best-Practices): Gewohnheiten, die typische Überraschungen vermeiden

**Wenn etwas nicht funktioniert**

- [Fehlerbehebung](Troubleshooting): Fehlermeldungen und Dokumente, die nicht richtig aussehen
- [FAQ](FAQ): kurze Antworten auf häufige Fragen
- [Sicherheit](Security): Probleme melden und mit nicht vertrauenswürdigen Eingaben umgehen

**Weiterführendes**

- [Architektur](Architecture): wie aus Ruby-Aufrufen ein `.docx`-Paket wird
- [Integrationen](Integrations): Rails-Controller, Rack-Anwendungen und Hintergrundjobs
- [Deployment](Deployment): docxi in Produktion betreiben, und wie Versionen veröffentlicht werden
- [Mitwirken](Contributing) und [Entwicklungsumgebung](Development-Setup)
- [Roadmap](Roadmap), [Änderungsprotokoll](Changelog) und [Danksagungen](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>Was docxi kann und was nicht

docxi erstellt neue Dokumente. Es kann keine vorhandene `.docx`-Datei öffnen oder bearbeiten, keine Word-Vorlage ausfüllen und keinen Seriendruck ausführen. Es gibt keine nummerierten Listen, keine Überschriftenformatvorlagen und kein automatisches Inhaltsverzeichnis, und die Seiten sind immer US Letter im Hochformat. Die für 0.2.0 geplanten bekannten Probleme stehen in der [Roadmap](Roadmap).

## <a id="credits"></a>Danksagungen

docxi wird von [Hammad Maqbool](https://github.com/hammad417) bei [Phaedra Solutions](https://www.phaedrasolutions.com/) gepflegt. Das Projekt entstand 2016 auf Basis von [docxer](https://github.com/gtimti/docxer) von Timur Gabdrakipov. Details stehen unter [Danksagungen](Credits).

## <a id="languages"></a>Sprachen

Die englische Fassung ist die Referenzversion dieses Wikis. Die Übersetzungen sind in der Seitenleiste und oben auf jeder Seite verlinkt.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Problem melden](https://github.com/hammad417/docxi/issues)
