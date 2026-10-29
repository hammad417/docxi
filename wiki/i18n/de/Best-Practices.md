<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# docxi Best Practices: typische Überraschungen vermeiden

Diese Gewohnheiten ergeben sich direkt aus der Funktionsweise von docxi, und die meisten verhindern ein Problem, das unter [Fehlerbehebung](Troubleshooting) beschrieben ist. Sie sind nach Themen gruppiert. Die Hinweise zur Seiteneinrichtung gelten für alle, die zu Sicherheit und Performance vor allem für Produktionsanwendungen.

## <a id="page-setup"></a>Seiteneinrichtung

- **Übergeben Sie alle vier Ränder zusammen.** Der Hash, den Sie an `Docxi::Document.new` übergeben, ersetzt die Standardwerte, und jeder der Werte `top`, `right`, `bottom` und `left`, den Sie weglassen, wird leer geschrieben.
- **Setzen Sie Kopf- und Fußzeile der ersten Seite.** docxi aktiviert immer eine eigene erste Seite. Wenn Seite 1 eine Kopf- oder Fußzeile braucht, rufen Sie zusätzlich `add_first_header` und `add_first_footer` auf.

## <a id="text"></a>Text

- **Verwenden Sie `space: true` für führende oder nachgestellte Leerzeichen.** Sonst entfernt Word sie, und die Wörter zweier Textläufe stoßen aneinander.
- **Legen Sie gemeinsame Formatierung auf den Absatz.** Optionen an `p`, etwa `size:` oder `font:`, gelten für jeden Textlauf darin, also müssen Sie weniger wiederholen.
- **Verwenden Sie `p.br` für Zeilenumbrüche und `document.br(page: true)` für Seitenumbrüche.** `document.br` ohne `page: true` erzeugt ungültiges XML.

## <a id="tables"></a>Tabellen

- **Setzen Sie Rahmen auf Zeilen oder Zellen, nicht auf die Tabelle.** In 0.1.0 wird `borders:` auf Tabellenebene auf jede Zelle angewendet statt auf den äußeren Rand der Tabelle.
- **Geben Sie `columns_width` für jede Spalte an.** Die Zellen übernehmen ihre Breite daraus, und die Tabellen behalten das erwartete Layout.

## <a id="images"></a>Bilder

- **Verwenden Sie PNG-Dateien.** Jedes Bild wird als PNG-Teil gespeichert.
- **Setzen Sie die Größe von Inline-Bildern mit `style:`**, zum Beispiel `style: { width: "120pt", height: "60pt" }`. Das ist die einzige Bildgröße, die in 0.1.0 wie erwartet funktioniert.
- **Übergeben Sie einen Pfad statt eines geöffneten `File`, wenn das Dokument mehrmals gerendert werden könnte.** IO-Bilder werden während `render` gelesen, und ein zweites Rendern erhält ein leeres Bild.
- **Bilder in Kopfzeilen, Tabellenzellen und positionierte Bilder bis 0.2.0.** Breite und Höhe werden beide als `width * height * 14.92` EMU berechnet, daher sind sie quadratisch. Für ein quadratisches Bild mit `S` Zoll Kantenlänge übergeben Sie `width` und `height` jeweils als `Math.sqrt(S * 914400 / 14.92).round`, für einen Zoll also etwa 248.

## <a id="security"></a>Sicherheit

- **Übergeben Sie keine nicht vertrauenswürdigen Strings an `add_media`.** Ein Pfad lässt docxi diese Datei von Ihrem Server lesen, und eine URL lässt Ihren Server sie herunterladen. Übergeben Sie ein IO, das Sie selbst geöffnet haben, oder prüfen Sie den Wert vorher. Siehe [Sicherheit](Security).
- **Text wird für Sie maskiert.** Strings, die an `text` übergeben werden, werden als XML-Text geschrieben, Benutzereingaben können also die Dokumentstruktur nicht verändern.

## <a id="performance-and-reliability"></a>Performance und Zuverlässigkeit

- **Erzeugen Sie große Dokumente in einem Hintergrundjob.** Das ganze Dokument wird beim Aufruf von `render` im Speicher aufgebaut. Siehe [Integrationen](Integrations).
- **Einmal rendern, dann die Bytes wiederverwenden.** Behalten Sie `word.render.string`, wenn Sie das Dokument mehrmals senden oder speichern müssen.
- **Legen Sie eine kompatible Version fest**, mit `gem "docxi", "~> 0.1"`.
- **Testen Sie, was Sie erzeugen.** Eine `.docx`-Datei ist ein Zip aus XML-Dateien, Tests können sie also entpacken und das XML prüfen. Die Tests von docxi selbst arbeiten so. Siehe [Architektur](Architecture).

## <a id="related-pages"></a>Verwandte Seiten

- [Funktionen](Features)
- [Konfiguration](Configuration)
- [Fehlerbehebung](Troubleshooting)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
