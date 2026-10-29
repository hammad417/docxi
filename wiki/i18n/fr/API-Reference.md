<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Référence de l'API de docxi : classes et méthodes

docxi est une bibliothèque Ruby : cette page liste donc son API Ruby. Il n'y a pas d'API HTTP. Les options de chaque méthode sont décrites dans [Fonctionnalités](Features), et les classes sont expliquées dans [Architecture](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Méthode | Renvoie | Remarques |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` correspond aux marges de la page. Voir [Configuration](Configuration) |
| `#document` | `Docxi::Word::Document` | Le corps auquel vous ajoutez du contenu |
| `#render` | `StringIO` | Le paquet `.docx` complet. `.string` donne les octets |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Méthode | Renvoie | Ajoute |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Un paragraphe |
| `text(string, options = {})` | Paragraph | Un paragraphe avec un seul segment de texte |
| `br(page: true)` | Break | Un saut de page |
| `tab(page: true)` | Tab | Un paragraphe contenant une tabulation |
| `table(options = {}) { \|table\| }` | Table | Un tableau |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Des entrées de table des matières |
| `frame(text, options = {})` | Paragraph | Un paragraphe contenant une zone de texte |
| `add_media(io_path_or_url, options = {})` | Media | Une image à placer avec `image` |
| `add_hyperlink(url, options = {})` | String | Un identifiant de relation pour `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | En-tête à partir de la page 2, et en-tête de la page 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Pied de page à partir de la page 2, et pied de page de la page 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Paragraphe (`Docxi::Word::Contents::Paragraph`)

| Méthode | Ajoute |
|---|---|
| `text(string, options = {})` | Un segment de texte. Les options du paragraphe sont ses valeurs par défaut |
| `br(times: 1)` | Des sauts de ligne |
| `tab(times: 1)` | Des caractères de tabulation |
| `image(media, options = {})` | Une image en ligne ou positionnée |
| `hyperlink(text, url, id)` | Un lien qui utilise l'identifiant de relation `id` renvoyé par `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Une zone de texte |

## <a id="tables"></a>Tableaux

| Méthode | Ajoute |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Une ligne. Les options du tableau sont ses valeurs par défaut |
| `TableRow#tc(options = {}) { \|tc\| }` | Une cellule. Les options de la ligne sont ses valeurs par défaut |
| `TableCell#text(string, options = {})` | Un paragraphe. Les options de la cellule sont ses valeurs par défaut |
| `TableCell#p(options = {}) { \|p\| }` | Un paragraphe avec ses propres options |
| `TableCell#image(media, width:, height:)` | Une image |
| `TableCell#page_numbers` | Un champ de numéro de page |

## <a id="table-of-contents"></a>Table des matières

| Méthode | Ajoute |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Une entrée de la table des matières |

## <a id="headers-and-footers"></a>En-têtes et pieds de page

| Classe | Méthodes |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Cadres

`Docxi::Word::Contents::Frame` accepte `text`, `br`, `tab`, `image` et `frame` dans son bloc.

## <a id="constants-and-helpers"></a>Constantes et méthodes utilitaires

| Nom | Signification |
|---|---|
| `Docxi::VERSION` | La version installée, par exemple `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Écrit une partie dans l'archive en fixant sa taille à l'avance. Utilisée en interne par chaque partie |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Méthodes utilitaires internes qui suivent les règles de `blank?` d'ActiveSupport |

## <a id="related-pages"></a>Pages associées

- [Fonctionnalités](Features)
- [Exemples d'utilisation](Usage-Examples)
- [Architecture](Architecture)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
