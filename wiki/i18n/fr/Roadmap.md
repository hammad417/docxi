<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Feuille de route et problèmes connus de docxi

docxi n'a ni calendrier de versions ni feuille de route avec des engagements. Le seul travail prévu et consigné consiste à corriger les problèmes connus ci-dessous. Ces corrections sont prévues pour la 0.2.0, sans date. Elles ont été tenues à l'écart de la 0.1.0 parce qu'elles modifient les documents que les utilisateurs actuels génèrent déjà.

## <a id="known-issues-planned-for-020"></a>Problèmes connus prévus pour la 0.2.0

| Problème | Contournement dans la 0.1.0 |
|---|---|
| Les images positionnées, d'en-tête et de cellule calculent leur taille comme `width * height * 14.92` EMU pour les deux dimensions | Utilisez des images en ligne dimensionnées avec `style:`, ou voir [Bonnes pratiques](Best-Practices#images) |
| Les `borders:` au niveau du tableau s'appliquent à chaque cellule au lieu du contour du tableau | Pour le contour seul, définissez `borders:` sur les côtés extérieurs des cellules du bord |
| Le code du champ de numéro de page est écrit sans la barre oblique inverse de `\* MERGEFORMAT` | Aucun |
| Les numéros de page dans les cellules de tableau sont précédés de « GlobalOptions » | Placez les numéros de page dans un en-tête ou un pied de page |
| Les images sont toujours stockées comme parties PNG, et `[Content_Types].xml` ne déclare que les deux premiers en-têtes et pieds de page | Utilisez des images PNG, et au plus deux en-têtes et deux pieds de page |
| Ne passer que certaines marges écrit les autres vides | Passez `top`, `right`, `bottom` et `left` ensemble |
| `docProps/core.xml` enregistre l'heure locale avec un marqueur UTC et n'écrit jamais l'auteur | Aucun |
| `Docxi.test` lève `NoMethodError` | Ne l'appelez pas. C'est une méthode utilitaire restée d'une ancienne version |
| `document.br` et `document.tab` sans `page: true` produisent du XML non valide | Utilisez `p.br` et `p.tab` dans un paragraphe |

## <a id="not-planned"></a>Non prévu

Ces points ne figurent pas dans la liste ci-dessus, et rien n'a été décidé à leur sujet :

- modifier des documents existants ou remplir des modèles
- listes numérotées, styles de titre ou table des matières automatique
- autres formats de page ou pages en paysage

## <a id="related-pages"></a>Pages associées

- [Journal des modifications](Changelog)
- [Contribuer](Contributing)
- [FAQ](FAQ)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
