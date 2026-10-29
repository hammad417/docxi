<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Erste Schritte mit docxi: Ihr erstes Word-Dokument in Ruby

Diese Anleitung führt Sie von einem leeren Ordner zu einer formatierten `.docx`-Datei. Sie setzt Ruby-Grundkenntnisse voraus. Rails oder Microsoft Word brauchen Sie nicht.

## <a id="1-install-the-gem"></a>1. Gem installieren

```sh
gem install docxi
```

In einem Projekt mit Gemfile fügen Sie stattdessen `gem "docxi", "~> 0.1"` hinzu und führen `bundle install` aus. Die Details stehen unter [Installation](Installation).

## <a id="2-create-a-document"></a>2. Dokument erstellen

Speichern Sie Folgendes als `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Führen Sie `ruby hello.rb` aus (oder `bundle exec ruby hello.rb` in einem Bundler-Projekt) und öffnen Sie `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Was jede Zeile macht

- `Docxi::Document.new` erstellt ein leeres Dokument mit den Standard-Seitenrändern.
- `word.document` ist der Hauptteil des Dokuments. Alles, was Sie hinzufügen, läuft darüber.
- `document.p` fügt einen Absatz hinzu. Optionen für `p`, etwa `align:`, gelten für den Absatz und den Text darin.
- `p.text` fügt einen Textlauf (Run) hinzu. `size:` wird in Punkt angegeben.
- `word.render` baut das `.docx`-Paket und gibt ein `StringIO` zurück. Dessen `.string` enthält die Bytes der Datei.

## <a id="4-add-a-list-and-a-table"></a>4. Liste und Tabelle hinzufügen

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

Zwei Dinge sind wichtig:

- `space: true` erhält das Leerzeichen am Anfang oder Ende eines Textlaufs. Ohne diese Option entfernt Word das Leerzeichen, und die Wörter laufen ineinander.
- Optionen an `tr` gelten für jede Zelle der Zeile. Deshalb ist die ganze Kopfzeile der Tabelle schattiert und fett.

## <a id="5-where-to-go-next"></a>5. Wie es weitergeht

- [Anwendungsbeispiele](Usage-Examples): vollständige Rezepte für Berichte, Rechnungen und Briefe
- [Funktionen](Features): alle Optionen für Absätze, Text, Tabellen, Bilder, Kopf- und Fußzeilen
- [Integrationen](Integrations): das Dokument aus einem Rails-Controller senden
- [Best Practices](Best-Practices): typische Überraschungen vermeiden, bevor sie auftreten

---

[← Wiki-Startseite](Home) · [README](https://github.com/hammad417/docxi#readme)
