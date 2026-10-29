<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Bonnes pratiques avec docxi : éviter les surprises courantes

Ces habitudes découlent directement du fonctionnement de docxi, et la plupart évitent un problème listé dans [Dépannage](Troubleshooting). Elles sont regroupées par thème. Les conseils de mise en page s'adressent à tout le monde, et ceux sur la sécurité et les performances comptent surtout pour les applications en production.

## <a id="page-setup"></a>Mise en page

- **Passez les quatre marges ensemble.** Le hash que vous passez à `Docxi::Document.new` remplace les valeurs par défaut, et chacune des valeurs `top`, `right`, `bottom` et `left` que vous omettez est écrite vide.
- **Définissez l'en-tête et le pied de page de la première page.** docxi active toujours une première page distincte. Si la page 1 a besoin d'un en-tête ou d'un pied de page, appelez aussi `add_first_header` et `add_first_footer`.

## <a id="text"></a>Texte

- **Utilisez `space: true` pour les espaces en début ou en fin de segment.** Sinon, Word les supprime et les mots de deux segments se collent.
- **Placez la mise en forme commune sur le paragraphe.** Les options passées à `p`, comme `size:` ou `font:`, s'appliquent à chaque segment qu'il contient, ce qui évite de les répéter.
- **Utilisez `p.br` pour les sauts de ligne et `document.br(page: true)` pour les sauts de page.** `document.br` sans `page: true` produit du XML non valide.

## <a id="tables"></a>Tableaux

- **Définissez les bordures sur les lignes ou les cellules, pas sur le tableau.** Dans la 0.1.0, les `borders:` au niveau du tableau s'appliquent à chaque cellule au lieu du contour du tableau.
- **Indiquez `columns_width` pour chaque colonne.** Les cellules en tirent leur largeur, et les tableaux gardent la disposition attendue.

## <a id="images"></a>Images

- **Utilisez des fichiers PNG.** Chaque image est stockée comme une partie PNG.
- **Dimensionnez les images en ligne avec `style:`**, par exemple `style: { width: "120pt", height: "60pt" }`. C'est le seul réglage de taille d'image qui fonctionne comme prévu dans la 0.1.0.
- **Passez un chemin plutôt qu'un `File` ouvert si le document peut être généré plusieurs fois.** Les images IO sont lues pendant `render`, et une deuxième génération obtient une image vide.
- **Images d'en-tête, de cellule et positionnées jusqu'à la 0.2.0.** Leur largeur et leur hauteur sont toutes deux calculées comme `width * height * 14.92` EMU, si bien qu'elles sont carrées. Pour obtenir une image carrée de `S` pouces de côté, passez `width` et `height` égaux à `Math.sqrt(S * 914400 / 14.92).round`, soit environ 248 pour un pouce.

## <a id="security"></a>Sécurité

- **Ne passez pas de chaînes non fiables à `add_media`.** Un chemin fait lire ce fichier sur votre serveur par docxi, et une URL fait télécharger la ressource par votre serveur. Passez un IO que vous avez ouvert vous-même, ou vérifiez d'abord la valeur. Voir [Sécurité](Security).
- **Le texte est échappé pour vous.** Les chaînes passées à `text` sont écrites comme texte XML : les saisies des utilisateurs ne peuvent donc pas modifier la structure du document.

## <a id="performance-and-reliability"></a>Performances et fiabilité

- **Générez les gros documents dans une tâche en arrière-plan.** Tout le document est construit en mémoire lorsque vous appelez `render`. Voir [Intégrations](Integrations).
- **Générez une fois, puis réutilisez les octets.** Conservez `word.render.string` si vous devez envoyer ou stocker le document plusieurs fois.
- **Fixez une version compatible** avec `gem "docxi", "~> 0.1"`.
- **Testez ce que vous générez.** Un `.docx` est un zip de fichiers XML : les tests peuvent le décompresser et vérifier le XML. Les tests de docxi fonctionnent ainsi. Voir [Architecture](Architecture).

## <a id="related-pages"></a>Pages associées

- [Fonctionnalités](Features)
- [Configuration](Configuration)
- [Dépannage](Troubleshooting)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
