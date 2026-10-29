<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# docxi konfigurieren: Seitenränder, Seiteneinrichtung und Einheiten

docxi hat keine Konfigurationsdatei, keine Umgebungsvariablen und keine globalen Einstellungen. Alles wird pro Dokument oder pro Element festgelegt. Diese Seite beschreibt die Einstellungen auf Dokumentebene und die Einheiten, die die Optionen verwenden. Die Optionen der einzelnen Elemente stehen unter [Funktionen](Features).

## <a id="page-margins"></a>Seitenränder

Die Seitenränder werden beim Erstellen des Dokuments übergeben:

```ruby
require "docxi"

# Alle Werte in Zwanzigstelpunkt (1440 = 1 Zoll).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Option | Standard | Bedeutung |
|---|---|---|
| `top` | `720` (0,5 Zoll) | Oberer Rand |
| `right` | `1440` (1 Zoll) | Rechter Rand |
| `bottom` | `720` (0,5 Zoll) | Unterer Rand |
| `left` | `1440` (1 Zoll) | Linker Rand |
| `header` | `190` | Abstand von der oberen Seitenkante zur Kopfzeile |
| `footer` | `200` | Abstand von der unteren Seitenkante zur Fußzeile |

Der übergebene Hash ersetzt die Standardwerte, setzen Sie `top`, `right`, `bottom` und `left` also gemeinsam. Jeder dieser vier Werte, den Sie weglassen, wird leer geschrieben. `header` und `footer` fallen auf 190 und 200 zurück, wenn sie fehlen oder leer sind.

## <a id="page-size-and-orientation"></a>Seitengröße und Ausrichtung

Die Seiten haben das Format US Letter im Hochformat: 12240 × 15840 Zwanzigstelpunkt, also 8,5 × 11 Zoll. Größe und Ausrichtung lassen sich in 0.1.0 nicht ändern.

## <a id="first-page"></a>Erste Seite

docxi aktiviert immer die Word-Einstellung „Erste Seite anders“. Kopf- und Fußzeilen, die mit `add_header` und `add_footer` hinzugefügt werden, gelten ab Seite 2. Seite 1 verwendet `add_first_header` und `add_first_footer`, setzen Sie diese also ebenfalls, wenn die erste Seite eine Kopf- oder Fußzeile braucht. Siehe [Funktionen](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Was jedes Dokument enthält

- Die Standardschrift ist Arial, 12 pt.
- Definierte Formatvorlagen: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` und `NoList`.
- Aufzählungspunkte (`ul: true`) verwenden einen Gedankenstrich als Aufzählungszeichen.
- Die Dokumenteigenschaften nennen als Anwendung „Docxi Ruby Gem“ und speichern den Zeitpunkt, zu dem das Dokument gerendert wurde.

## <a id="units"></a>Einheiten

| Option | Einheit |
|---|---|
| Text-`size:` | Punkt (docxi verdoppelt den Wert zu den halben Punkten von Word) |
| Seitenränder, Absatz-`left:` und `right:`, Tabellen-`iwidth:`, Zeilen-`height:` | Zwanzigstelpunkt (1440 pro Zoll) |
| Größen von Zellen-`borders:` | Achtelpunkt |
| `columns_width:` und Zellen-`width:` | Werden mit 14.92 multipliziert, um Zwanzigstelpunkt zu erhalten, ungefähr Pixel bei 96 dpi |
| `style:` von Inline-Bildern | CSS-Längen wie `"120pt"` |
| `pH:` und `pV:` von positionierten Bildern | EMU (914400 pro Zoll) |
| Text- und Füllfarben | Hex ohne `#`, zum Beispiel `"8B0000"` |
| `fill_color:` von Bildern | VML-Farbe mit `#`, zum Beispiel `"#00FF00"` |

## <a id="related-pages"></a>Verwandte Seiten

- [Funktionen](Features)
- [API-Referenz](API-Reference)
- [Best Practices](Best-Practices)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
