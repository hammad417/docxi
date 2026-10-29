<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Primi passi con docxi: il tuo primo documento Word in Ruby

Questa guida ti porta da una cartella vuota a un file `.docx` formattato. Presuppone che tu conosca le basi di Ruby. Non ti servono né Rails né Microsoft Word.

## <a id="1-install-the-gem"></a>1. Installa la gem

```sh
gem install docxi
```

In un progetto con un Gemfile, aggiungi invece `gem "docxi", "~> 0.1"` ed esegui `bundle install`. Trovi i dettagli in [Installazione](Installation).

## <a id="2-create-a-document"></a>2. Crea un documento

Salva questo codice come `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Esegui `ruby hello.rb` (oppure `bundle exec ruby hello.rb` in un progetto Bundler) e apri `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Cosa fa ogni riga

- `Docxi::Document.new` crea un documento vuoto con i margini di pagina predefiniti.
- `word.document` è il corpo del documento. Tutto ciò che aggiungi passa da lì.
- `document.p` aggiunge un paragrafo. Le opzioni passate a `p`, come `align:`, si applicano al paragrafo e al testo al suo interno.
- `p.text` aggiunge un frammento di testo (run). `size:` è espresso in punti.
- `word.render` costruisce il pacchetto `.docx` e restituisce uno `StringIO`. Il suo `.string` contiene i byte del file.

## <a id="4-add-a-list-and-a-table"></a>4. Aggiungi un elenco e una tabella

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

Due cose da notare:

- `space: true` conserva lo spazio all'inizio o alla fine di un frammento. Senza questa opzione Word elimina lo spazio e le parole si attaccano.
- Le opzioni di `tr` si applicano a ogni cella della riga, quindi l'intera riga di intestazione risulta ombreggiata e in grassetto.

## <a id="5-where-to-go-next"></a>5. Dove andare dopo

- [Esempi d'uso](Usage-Examples): ricette complete per report, fatture e lettere
- [Funzionalità](Features): tutte le opzioni per paragrafi, testo, tabelle, immagini, intestazioni e piè di pagina
- [Integrazioni](Integrations): inviare il documento da un controller Rails
- [Buone pratiche](Best-Practices): evitare in anticipo le sorprese più comuni

---

[← Home della wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
