<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Fonctionnalités de docxi : paragraphes, texte, tableaux, images, en-têtes et pieds de page

Cette page liste chaque élément que docxi peut ajouter à un document Word, avec les options que chacun accepte. Les exemples partent de cette configuration :

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Les options dont vous n'avez pas besoin peuvent être omises. Les couleurs sont des valeurs hexadécimales sans `#`, par exemple `"8B0000"`. Pour les unités et la mise en page, voir [Configuration](Configuration).

## <a id="paragraphs"></a>Paragraphes

`document.p(options) { |p| ... }` ajoute un paragraphe.

| Option | Valeurs | Effet |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Alignement. `"both"` justifie le texte |
| `fill` | Couleur hexadécimale | Trame de fond derrière le paragraphe et son texte |
| `left`, `right` | Vingtièmes de point | Retrait. Définissez les deux, car un seul est ignoré |
| `ul` | `true` | Fait du paragraphe un élément de liste à puces |
| `bottom` | `true` | Trace une bordure sous le paragraphe |

Les options de texte comme `size:`, `bold:` ou `color:` peuvent aussi être passées à `p`. Elles deviennent les valeurs par défaut de chaque segment de texte du paragraphe.

Dans le bloc, vous pouvez appeler `text`, `br`, `tab`, `image`, `hyperlink` et `frame`. `document.text(string, options)` est un raccourci pour un paragraphe qui contient un seul segment.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Texte

`p.text(string, options)` ajoute un segment de texte (run).

| Option | Valeurs | Effet |
|---|---|---|
| `bold` | `true` | Gras |
| `italic` | `true` | Italique |
| `underline` | `true` | Soulignement simple |
| `color` | Couleur hexadécimale | Couleur du texte |
| `size` | Points | Taille de police |
| `font` | Nom de police | Police, par exemple `"Georgia"` |
| `fill` | Couleur hexadécimale | Trame de fond derrière le texte |
| `space` | `true` | Conserve les espaces au début et à la fin du texte |
| `br` | `true` | Ajoute un saut de ligne après le texte |

Une chaîne `nil` est ignorée.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Sauts de ligne, sauts de page et tabulations

- `p.br`, ou `p.br(times: 3)`, ajoute des sauts de ligne dans un paragraphe.
- `document.br(page: true)` commence une nouvelle page.
- `p.tab`, ou `p.tab(times: 2)`, ajoute des caractères de tabulation dans un paragraphe.
- `document.tab(page: true)` ajoute un paragraphe qui contient une seule tabulation.

N'appelez pas `document.br` ni `document.tab` sans `page: true`. Ils placent un segment de texte directement dans le corps du document, ce qui n'est pas du WordprocessingML valide.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tableaux

`document.table(options)` prend un bloc. À l'intérieur, `table.tr(options)` ajoute des lignes et `tr.tc(options)` ajoute des cellules.

**Options du tableau**

| Option | Valeurs | Effet |
|---|---|---|
| `columns_width` | Array de nombres | Largeurs des colonnes, multipliées par 14.92 pour obtenir des vingtièmes de point. Ce sont aussi les largeurs par défaut des cellules |
| `height` | Vingtièmes de point | Hauteur minimale de chaque ligne |
| `iwidth` | Vingtièmes de point | Retrait du tableau par rapport à la marge gauche |
| `style` | Identifiant de style | Style du tableau. docxi définit `TableNormal` et `TableCustom` |
| `width` | Nombre | Largeur préférée du tableau, écrite avec le type `auto`. Sans `columns_width`, chaque cellule reçoit aussi cette largeur, multipliée par 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Alignement du tableau. Le texte ajouté avec `tc.text` est aligné de la même façon |
| `padding` | `true` | Ajoute une marge intérieure fixe dans les cellules |
| `borders` | Hash | Problème connu : s'applique à chaque cellule au lieu du contour du tableau |

**Options de ligne** : `height:` définit la hauteur minimale de la ligne. Toute option de cellule passée à `tr` s'applique à chaque cellule de la ligne. Les options du tableau sont transmises de la même façon à chaque ligne et à chaque cellule.

**Options de cellule**

| Option | Valeurs | Effet |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Épaisseur de bordure en huitièmes de point. `nil` signifie pas de bordure de ce côté |
| `fill` | Couleur hexadécimale | Trame de fond de la cellule. Les cellules sans remplissage sont blanches (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Alignement vertical |
| `colspan` | Entier | Nombre de colonnes couvertes par la cellule |
| `rowspan` | `true` | Commence une fusion verticale |
| `merged` | `true` | Prolonge la fusion verticale de la cellule du dessus. Son contenu est ignoré |
| `width` | Nombre | Largeur de la cellule, multipliée par 14.92 |

**Contenu des cellules**

- `tc.text(string, options)` : les options de cellule et de ligne, comme `bold:` ou `align:`, s'appliquent aussi au texte.
- `tc.p(options) { |p| ... }` : un paragraphe qui n'utilise que ses propres options.
- `tc.image(media, width:, height:)` : une image. Voir le problème de taille dans [Images](#images).
- `tc.page_numbers` : un champ de numéro de page. Problème connu : il est précédé du texte « GlobalOptions ».

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Images

`media = document.add_media(source)` enregistre une image. `source` peut être :

- un objet IO, comme un `File` ou un `StringIO`
- un chemin de fichier
- une URL `http`, `https` ou `ftp`

L'image est lue lorsque vous appelez `render` et toujours stockée sous forme de partie PNG : les fichiers PNG sont donc le choix sûr.

**Images en ligne** : `p.image(media, options)` place l'image dans la ligne de texte.

| Option | Valeurs | Effet |
|---|---|---|
| `style` | Hash de propriétés CSS | Taille, par exemple `{ width: "120pt", height: "60pt" }` |
| `title` | Chaîne | Titre de l'image |
| `fill_color` | Couleur VML comme `"#00FF00"` | Remplissage derrière l'image |
| `opacity` | Chaîne comme `"50%"` | Opacité du remplissage |
| `wrap` | `true` | Désactive l'habillage du texte autour de l'image |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Images positionnées** : `p.image(media, pH:, pV:, wrap: true, width:, height:)` ancre l'image avec un décalage par rapport à la colonne (`pH`) et au paragraphe (`pV`), en EMU. `pH`, `pV` et `wrap` doivent tous être fournis, sinon l'image est placée en ligne.

**Images d'en-tête** : `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` désactive l'habillage. Toute autre valeur fait habiller le texte autour de l'image en mode carré.

**Problème connu dans la 0.1.0** : les images positionnées, les images d'en-tête et les images de cellule calculent leur largeur et leur hauteur comme `width * height * 14.92` EMU. Elles sont toujours carrées et généralement minuscules. Les images en ligne dimensionnées avec `style:` ne sont pas concernées. Voir [Bonnes pratiques](Best-Practices#images) pour une solution de contournement.

## <a id="headers-and-footers"></a>En-têtes et pieds de page

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Un en-tête peut contenir `text`, `br`, `table`, `image` et `page_numbers`. Un pied de page peut contenir `text`, `br`, `table` et `page_numbers`.
- Les options passées à `Header.new` ou `Footer.new` s'appliquent à leur texte.
- `add_header` et `add_footer` s'appliquent à partir de la page 2. La page 1 utilise toujours `add_first_header` et `add_first_footer`. Voir [Configuration](Configuration#first-page).

## <a id="page-numbers"></a>Numéros de page

`page_numbers(name: "Page", align: "center")` écrit le texte de `name` suivi d'un champ PAGE. Le numéro est en Arial 10 pt, gris foncé. `align` vaut `"right"` par défaut.

Problème connu : le code de champ est écrit sans la barre oblique inverse de `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Table des matières

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Avec un numéro de page, l'entrée reçoit un taquet de tabulation aligné à droite à la position `pos` (en vingtièmes de point ; 9360 correspond à 6,5 pouces), rempli avec le caractère `leader`. Les valeurs des caractères de suite viennent de WordprocessingML, par exemple `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` et `"none"`.
- Sans numéro de page, l'entrée est le texte placé après une tabulation.
- Les numéros de page sont ceux que vous passez. docxi ne crée pas de champ TOC Word et ne calcule pas les pages.

## <a id="hyperlinks"></a>Liens hypertexte

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` renvoie un identifiant de relation, et `hyperlink(text, url, id)` place le lien. La cible vient de la relation : passez donc l'identifiant renvoyé pour l'URL voulue. Seuls les liens externes sont pris en charge.

## <a id="frames-text-boxes"></a>Cadres (zones de texte)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Les cadres sont des zones de texte VML avec un remplissage fixe gris clair semi-transparent et un contour noir. Le texte s'organise au-dessus et en dessous.
- Dans un cadre, vous pouvez appeler `text`, `br`, `tab`, `image` et `frame`.
- `spacing: true` ajoute de l'espace avant le texte.

## <a id="output"></a>Sortie

`word.render` renvoie un `StringIO` contenant le paquet `.docx` complet. L'archive est écrite sans enregistrements Zip64, si bien que Google Docs et les anciennes versions de LibreOffice peuvent l'ouvrir. La page [Architecture](Architecture) liste chaque partie du paquet.

## <a id="related-pages"></a>Pages associées

- [Exemples d'utilisation](Usage-Examples)
- [Référence de l'API](API-Reference)
- [Feuille de route](Roadmap), pour les problèmes connus prévus pour la 0.2.0

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
