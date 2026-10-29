<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Fehlerbehebung: Installationsfehler, Laufzeitfehler und fehlerhafte Ausgabe

Suchen Sie unten Ihr Symptom. Die meisten Probleme in älteren Anwendungen stammen von docxi 0.0.x, und ein Upgrade auf 0.1.0 behebt sie ohne Codeänderungen.

## <a id="installation"></a>Installation

**Bundler findet in einer Rails-6+-Anwendung keine kompatiblen Versionen von `nokogiri`.**
docxi 0.0.6 legt `nokogiri < 1.7` fest, während jede Rails-Version seit 6.0 nokogiri 1.8.5 oder neuer braucht. Verwenden Sie `gem "docxi", "~> 0.1"`.

**Bundler findet keine kompatiblen Versionen von `rubyzip`.**
docxi 0.0.6 legt `rubyzip ~> 0.9.9` fest, was mit jedem Gem kollidiert, das ein neueres rubyzip braucht. docxi 0.1.0 akzeptiert rubyzip 1.3 bis 3.x.

**Bundler installiert weiterhin 0.0.6.**
Entweder legt das Gemfile docxi auf eine 0.0.x-Version fest, zum Beispiel mit `"~> 0.0.6"`, oder Ruby ist älter als 2.5, das 0.1.0 voraussetzt. Ändern Sie die Zeile im Gemfile in `gem "docxi", "~> 0.1"` und prüfen Sie `ruby -v`.

**nokogiri lässt sich nicht kompilieren.**
Für Ihre Plattform gibt es kein passendes vorkompiliertes nokogiri-Gem, daher wird es aus dem Quellcode gebaut. Installieren Sie einen C-Compiler und folgen Sie der [Installationsanleitung von nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Fehler beim Erstellen eines Dokuments

**`NoMethodError: undefined method 'present?'` außerhalb von Rails.**
docxi 0.0.6 setzte ein geladenes ActiveSupport voraus. 0.1.0 braucht es nicht.

**`NameError` beim Aufruf von `tab`, `frame` oder `page_numbers` einer Kopfzeile.**
Diese Methoden waren in 0.0.6 defekt und funktionieren in 0.1.0, einschließlich `document.frame("text")`.

**`Errno::ENOENT` beim Hinzufügen eines Bildes.**
Der Pfad existiert nicht, relativ zu dem Verzeichnis, in dem Ihr Programm läuft. Unter Ruby 3 mit docxi 0.0.6 schlagen auch Bild-URLs auf diese Weise fehl, weil 0.0.6 `Kernel#open` verwendete. 0.1.0 lädt URLs mit `URI#open` herunter.

## <a id="documents-that-dont-look-right"></a>Dokumente, die nicht richtig aussehen

**Google Docs oder LibreOffice öffnet die Datei nicht.**
docxi 0.0.6 schreibt mit rubyzip 3 in jeden Eintrag Zip64-Einträge, und diese Programme lehnen sie ab. 0.1.0 schreibt Archive ohne sie.

**Die erste Seite hat keine Kopf- oder Fußzeile.**
docxi gibt der ersten Seite eine eigene Kopf- und Fußzeile. Setzen Sie sie mit `add_first_header` und `add_first_footer`. Siehe [Konfiguration](Configuration#first-page).

**Ein Rand fehlt, oder das Seitenlayout sieht falsch aus.**
Wenn Sie Ränder setzen, übergeben Sie `top`, `right`, `bottom` und `left` zusammen. Die weggelassenen werden leer geschrieben.

**Leerzeichen zwischen Textläufen verschwinden.**
Word entfernt Leerzeichen am Anfang oder Ende eines Textlaufs. Ergänzen Sie diesen `text`-Aufruf um `space: true`.

**Bilder in Kopfzeilen, in Tabellenzellen oder mit `pH`/`pV` sind winzig oder quadratisch.**
Bekanntes Problem in 0.1.0: Ihre Größe wird für beide Dimensionen als `width * height * 14.92` EMU berechnet. Verwenden Sie Inline-Bilder, deren Größe über `style:` gesetzt wird, oder den Workaround unter [Best Practices](Best-Practices#images).

**Ein Bild fehlt, wenn dasselbe Dokument zum zweiten Mal gerendert wird.**
Bilder aus einem geöffneten `File` oder einem anderen IO werden während `render` gelesen, und ein zweites Rendern liest einen leeren Stream. Übergeben Sie stattdessen einen Pfad oder rendern Sie nur einmal.

**Auf `table` gesetzte Rahmen erscheinen um jede Zelle.**
Bekanntes Problem: `borders:` auf Tabellenebene wird auf jede Zelle angewendet statt auf den äußeren Rand der Tabelle. Für einen reinen Außenrahmen setzen Sie `borders:` auf die Randzellen, und zwar nur für ihre äußeren Seiten.

**Seitenzahlen in Tabellenzellen beginnen mit „GlobalOptions“.**
Bekanntes Problem in `TableCell#page_numbers`. Setzen Sie Seitenzahlen stattdessen in eine Fuß- oder Kopfzeile.

**Ein JPEG- oder GIF-Bild sieht falsch aus oder erscheint nicht.**
Jedes Bild wird unabhängig vom tatsächlichen Format als PNG-Teil mit PNG-Inhaltstyp gespeichert, daher zeigen manche Programme andere Formate möglicherweise nicht an. Konvertieren Sie Bilder vor dem Hinzufügen nach PNG.

## <a id="inspecting-a-generated-file"></a>Eine erzeugte Datei untersuchen

Eine `.docx`-Datei ist ein Zip-Archiv, Sie können also hineinschauen:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Wenn das Problem bestehen bleibt, [eröffnen Sie ein Issue](https://github.com/hammad417/docxi/issues) mit einem kurzen Skript, das es reproduziert, und den Versionen von Ruby, docxi, nokogiri und rubyzip.

## <a id="related-pages"></a>Verwandte Seiten

- [FAQ](FAQ)
- [Best Practices](Best-Practices)
- [Installation](Installation)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
