<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-FAQ: Word-Dokumente in Ruby und Rails

Kurze Antworten auf die häufigsten Fragen. Jede Antwort verlinkt auf eine ausführlichere Erklärung.

**Muss Microsoft Word oder LibreOffice installiert sein?**
Nein. docxi schreibt das XML und das Zip-Archiv selbst.

**Braucht docxi Rails?**
Nein. Es funktioniert in jedem Ruby-Programm. Die Tests laufen außerdem mit geladenem Rails 5.2 bis 8.1. Siehe [Integrationen](Integrations).

**Wie erzeuge ich ein Word-Dokument in Ruby on Rails?**
Bauen Sie es in einem Controller oder einem Hintergrundjob und senden Sie `word.render.string` mit `send_data`. Siehe [Integrationen](Integrations).

**Welche Ruby- und Rails-Versionen werden unterstützt?**
Ruby 2.5 und neuer. Rails ist optional, und die CI testet Rails 5.2 bis 8.1. Siehe [Installation](Installation).

**Kann docxi eine vorhandene .docx-Datei bearbeiten oder eine Word-Vorlage ausfüllen?**
Nein, es erstellt nur neue Dokumente. Für Vorlagen mit Seriendruckfeldern verwenden Sie ein vorlagenbasiertes Gem wie sablon.

**Wird das Inhaltsverzeichnis aus den Überschriften erzeugt?**
Nein. Sie fügen jeden Eintrag und seine Seitenzahl selbst hinzu. Siehe [Funktionen](Features#table-of-contents).

**Kann ich nummerierte Listen oder Überschriftenformatvorlagen verwenden?**
Nicht in 0.1.0. Absätze können Aufzählungspunkte sein (`ul: true`), und Überschriften können Sie mit Textoptionen wie `size:` und `bold:` formatieren.

**Kann ich Seitengröße oder Ausrichtung ändern?**
Noch nicht. Die Seiten sind US Letter im Hochformat, und nur die Ränder lassen sich ändern. Siehe [Konfiguration](Configuration).

**Welche Bildformate funktionieren?**
PNG. Jedes Bild wird als PNG-Teil gespeichert.

**Kann ich ein Bild über eine URL hinzufügen?**
Ja. `add_media` akzeptiert `http`-, `https`- und `ftp`-URLs und lädt sie beim Aufruf von `render` herunter. Übergeben Sie keine URLs von Benutzern, ohne sie zu prüfen. Siehe [Sicherheit](Security).

**Kann ich das Dokument direkt in eine Datei schreiben?**
`render` gibt ein `StringIO` zurück. Schreiben Sie dessen `.string` mit `File.binwrite("out.docx", word.render.string)`.

**Kann docxi sehr große Dokumente streamen?**
Nein. Das ganze Dokument wird beim Aufruf von `render` im Speicher aufgebaut. Siehe [Deployment](Deployment).

**Kann ich Absätze von rechts nach links setzen?**
In 0.1.0 gibt es keine Rechts-nach-links-Optionen. Text wird so geschrieben, wie Sie ihn übergeben.

**Wo sind die bekannten Probleme aufgelistet?**
In der [Roadmap](Roadmap) und im README-Abschnitt „Limitations and known issues“ (Einschränkungen und bekannte Probleme).

## <a id="related-pages"></a>Verwandte Seiten

- [Fehlerbehebung](Troubleshooting)
- [Funktionen](Features)
- [Erste Schritte](Getting-Started)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
