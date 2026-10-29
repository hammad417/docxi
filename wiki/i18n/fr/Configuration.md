<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Configurer docxi : marges, mise en page et unités

docxi n'a ni fichier de configuration, ni variables d'environnement, ni réglages globaux. Tout se règle par document ou par élément. Cette page présente les réglages au niveau du document et les unités utilisées par les options. Les options de chaque élément sont décrites dans [Fonctionnalités](Features).

## <a id="page-margins"></a>Marges de la page

Les marges sont passées à la création du document :

```ruby
require "docxi"

# Toutes les valeurs sont en vingtièmes de point (1440 = 1 pouce).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Option | Valeur par défaut | Signification |
|---|---|---|
| `top` | `720` (0,5 pouce) | Marge supérieure |
| `right` | `1440` (1 pouce) | Marge droite |
| `bottom` | `720` (0,5 pouce) | Marge inférieure |
| `left` | `1440` (1 pouce) | Marge gauche |
| `header` | `190` | Distance entre le bord supérieur de la page et l'en-tête |
| `footer` | `200` | Distance entre le bord inférieur de la page et le pied de page |

Le hash que vous passez remplace les valeurs par défaut : définissez donc `top`, `right`, `bottom` et `left` ensemble. Chacune de ces quatre marges que vous omettez est écrite avec une valeur vide. `header` et `footer` reprennent les valeurs 190 et 200 lorsqu'ils sont absents ou vides.

## <a id="page-size-and-orientation"></a>Taille et orientation de la page

Les pages sont au format US Letter en orientation portrait : 12240 × 15840 vingtièmes de point, soit 8,5 × 11 pouces. La taille et l'orientation ne peuvent pas être modifiées dans la 0.1.0.

## <a id="first-page"></a>Première page

docxi active toujours l'option « Première page différente » de Word. Les en-têtes et pieds de page ajoutés avec `add_header` et `add_footer` s'appliquent à partir de la page 2. La page 1 utilise `add_first_header` et `add_first_footer` : définissez-les aussi si la première page a besoin d'un en-tête ou d'un pied de page. Voir [Fonctionnalités](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Ce que contient chaque document

- La police par défaut est Arial 12 pt.
- Styles définis : `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` et `NoList`.
- Les éléments de liste à puces (`ul: true`) utilisent un tiret comme puce.
- Les propriétés du document indiquent « Docxi Ruby Gem » comme nom d'application et enregistrent l'heure de génération du document.

## <a id="units"></a>Unités

| Option | Unité |
|---|---|
| `size:` du texte | Points (docxi les double pour obtenir les demi-points de Word) |
| Marges, `left:` et `right:` de paragraphe, `iwidth:` de tableau, `height:` de ligne | Vingtièmes de point (1440 par pouce) |
| Épaisseurs de `borders:` de cellule | Huitièmes de point |
| `columns_width:` et `width:` de cellule | Multipliés par 14.92 pour obtenir des vingtièmes de point, proche des pixels à 96 ppp |
| `style:` d'image en ligne | Longueurs CSS comme `"120pt"` |
| `pH:` et `pV:` d'image positionnée | EMU (914400 par pouce) |
| Couleurs du texte et du remplissage | Hexadécimal sans `#`, par exemple `"8B0000"` |
| `fill_color:` d'image | Couleur VML avec `#`, par exemple `"#00FF00"` |

## <a id="related-pages"></a>Pages associées

- [Fonctionnalités](Features)
- [Référence de l'API](API-Reference)
- [Bonnes pratiques](Best-Practices)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
