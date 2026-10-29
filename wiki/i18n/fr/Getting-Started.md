<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Premiers pas avec docxi : votre premier document Word en Ruby

Ce guide vous mène d'un dossier vide à un fichier `.docx` mis en forme. Il suppose que vous connaissez les bases de Ruby. Vous n'avez besoin ni de Rails ni de Microsoft Word.

## <a id="1-install-the-gem"></a>1. Installer la gem

```sh
gem install docxi
```

Dans un projet avec un Gemfile, ajoutez plutôt `gem "docxi", "~> 0.1"` et lancez `bundle install`. La page [Installation](Installation) donne tous les détails.

## <a id="2-create-a-document"></a>2. Créer un document

Enregistrez ceci sous le nom `hello.rb` :

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Lancez `ruby hello.rb` (ou `bundle exec ruby hello.rb` dans un projet Bundler) et ouvrez `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Le rôle de chaque ligne

- `Docxi::Document.new` crée un document vide avec les marges par défaut.
- `word.document` est le corps du document. Tout ce que vous ajoutez passe par lui.
- `document.p` ajoute un paragraphe. Les options passées à `p`, comme `align:`, s'appliquent au paragraphe et au texte qu'il contient.
- `p.text` ajoute un segment de texte (un run). `size:` est exprimé en points.
- `word.render` construit le paquet `.docx` et renvoie un `StringIO`. Son `.string` contient les octets du fichier.

## <a id="4-add-a-list-and-a-table"></a>4. Ajouter une liste et un tableau

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

Deux points à noter :

- `space: true` conserve l'espace au début ou à la fin d'un segment. Sans cette option, Word supprime l'espace et les mots se collent.
- Les options de `tr` s'appliquent à chaque cellule de la ligne : toute la ligne d'en-tête reçoit donc une trame de fond et passe en gras.

## <a id="5-where-to-go-next"></a>5. Pour aller plus loin

- [Exemples d'utilisation](Usage-Examples) : recettes complètes pour des rapports, des factures et des lettres
- [Fonctionnalités](Features) : toutes les options des paragraphes, du texte, des tableaux, des images, des en-têtes et des pieds de page
- [Intégrations](Integrations) : envoyer le document depuis un contrôleur Rails
- [Bonnes pratiques](Best-Practices) : éviter les surprises courantes avant qu'elles n'arrivent

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
