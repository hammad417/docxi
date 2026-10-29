<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi-Roadmap und bekannte Probleme

docxi hat keinen Veröffentlichungsplan und keine verbindliche Roadmap. Die einzige festgehaltene geplante Arbeit ist die Behebung der unten genannten bekannten Probleme. Sie ist für 0.2.0 geplant, ohne Termin. Die Korrekturen wurden aus 0.1.0 herausgehalten, weil sie die Dokumente verändern, die bestehende Nutzer bereits erzeugen.

## <a id="known-issues-planned-for-020"></a>Für 0.2.0 geplante bekannte Probleme

| Problem | Workaround in 0.1.0 |
|---|---|
| Positionierte Bilder, Bilder in Kopfzeilen und Bilder in Tabellenzellen berechnen ihre Größe für beide Dimensionen als `width * height * 14.92` EMU | Inline-Bilder verwenden, deren Größe über `style:` gesetzt wird, oder siehe [Best Practices](Best-Practices#images) |
| `borders:` auf Tabellenebene wird auf jede Zelle angewendet statt auf den äußeren Rand der Tabelle | Für einen reinen Außenrahmen `borders:` an den äußeren Seiten der Randzellen setzen |
| In der Feldfunktion der Seitenzahl fehlt der Backslash von `\* MERGEFORMAT` | Keiner |
| Seitenzahlen in Tabellenzellen wird „GlobalOptions“ vorangestellt | Seitenzahlen in eine Kopf- oder Fußzeile setzen |
| Bilder werden immer als PNG-Teile gespeichert, und `[Content_Types].xml` deklariert nur die ersten beiden Kopf- und Fußzeilen | PNG-Bilder und höchstens zwei Kopf- und zwei Fußzeilen verwenden |
| Werden nur einige Ränder übergeben, werden die übrigen leer geschrieben | `top`, `right`, `bottom` und `left` zusammen übergeben |
| `docProps/core.xml` speichert die Ortszeit mit UTC-Kennzeichnung und schreibt nie den Autor | Keiner |
| `Docxi.test` löst `NoMethodError` aus | Nicht aufrufen. Es ist eine übrig gebliebene Hilfsmethode |
| `document.br` und `document.tab` ohne `page: true` erzeugen ungültiges XML | `p.br` und `p.tab` innerhalb eines Absatzes verwenden |

## <a id="not-planned"></a>Nicht geplant

Diese Punkte stehen nicht auf der Liste oben, und es wurde nichts dazu entschieden:

- vorhandene Dokumente bearbeiten oder Vorlagen ausfüllen
- nummerierte Listen, Überschriftenformatvorlagen oder ein automatisches Inhaltsverzeichnis
- andere Seitengrößen oder Seiten im Querformat

## <a id="related-pages"></a>Verwandte Seiten

- [Änderungsprotokoll](Changelog)
- [Mitwirken](Contributing)
- [FAQ](FAQ)

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
