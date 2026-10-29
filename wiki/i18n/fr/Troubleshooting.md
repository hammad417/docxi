<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Dépannage de docxi : erreurs d'installation, erreurs d'exécution et documents défectueux

Trouvez votre symptôme ci-dessous. La plupart des problèmes rencontrés dans les applications anciennes viennent de docxi 0.0.x, et passer à la 0.1.0 les règle sans modifier le code.

## <a id="installation"></a>Installation

**Bundler ne trouve pas de versions compatibles de `nokogiri` dans une application Rails 6+.**
docxi 0.0.6 impose `nokogiri < 1.7`, alors que toutes les versions de Rails depuis la 6.0 exigent nokogiri 1.8.5 ou plus récent. Utilisez `gem "docxi", "~> 0.1"`.

**Bundler ne trouve pas de versions compatibles de `rubyzip`.**
docxi 0.0.6 impose `rubyzip ~> 0.9.9`, ce qui entre en conflit avec toute gem qui a besoin d'un rubyzip plus récent. docxi 0.1.0 accepte rubyzip 1.3 à 3.x.

**Bundler continue d'installer la 0.0.6.**
Soit le Gemfile fixe docxi sur une 0.0.x, par exemple avec `"~> 0.0.6"`, soit Ruby est antérieur à la version 2.5, que la 0.1.0 exige. Remplacez la ligne du Gemfile par `gem "docxi", "~> 0.1"` et vérifiez `ruby -v`.

**La compilation de nokogiri échoue.**
Aucune gem nokogiri précompilée ne correspond à votre plateforme : elle est donc compilée depuis les sources. Installez un compilateur C et suivez le [guide d'installation de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Erreurs lors de la construction d'un document

**`NoMethodError: undefined method 'present?'` en dehors de Rails.**
docxi 0.0.6 supposait qu'ActiveSupport était chargé. La 0.1.0 n'en a pas besoin.

**`NameError` lors de l'appel à `tab`, `frame` ou au `page_numbers` d'un en-tête.**
Ces méthodes ne fonctionnaient pas dans la 0.0.6 et fonctionnent dans la 0.1.0, y compris `document.frame("text")`.

**`Errno::ENOENT` lors de l'ajout d'une image.**
Le chemin n'existe pas par rapport au répertoire dans lequel votre programme s'exécute. Sous Ruby 3 avec docxi 0.0.6, les URL d'images échouent aussi de cette façon, car la 0.0.6 utilisait `Kernel#open`. La 0.1.0 télécharge les URL avec `URI#open`.

## <a id="documents-that-dont-look-right"></a>Documents qui ne s'affichent pas correctement

**Google Docs ou LibreOffice refuse d'ouvrir le fichier.**
docxi 0.0.6 avec rubyzip 3 écrit des enregistrements Zip64 dans chaque entrée, et ces programmes les rejettent. La 0.1.0 écrit les archives sans ces enregistrements.

**La première page n'a ni en-tête ni pied de page.**
docxi donne à la première page son propre en-tête et son propre pied de page. Définissez-les avec `add_first_header` et `add_first_footer`. Voir [Configuration](Configuration#first-page).

**Une marge manque, ou la mise en page semble incorrecte.**
Quand vous définissez les marges, passez `top`, `right`, `bottom` et `left` ensemble. Celles que vous omettez sont écrites vides.

**Les espaces entre les segments de texte disparaissent.**
Word supprime les espaces au début ou à la fin d'un segment. Ajoutez `space: true` à cet appel à `text`.

**Les images dans les en-têtes, dans les cellules de tableau ou avec `pH`/`pV` sont minuscules ou carrées.**
Problème connu dans la 0.1.0 : leur taille est calculée comme `width * height * 14.92` EMU pour les deux dimensions. Utilisez des images en ligne dimensionnées avec `style:`, ou consultez la solution de contournement dans [Bonnes pratiques](Best-Practices#images).

**Une image manque la deuxième fois que le même document est généré.**
Les images ajoutées depuis un `File` ouvert ou un autre IO sont lues pendant `render`, et une deuxième génération lit un flux vide. Passez plutôt un chemin, ou ne générez le document qu'une fois.

**Les bordures définies sur `table` apparaissent autour de chaque cellule.**
Problème connu : les `borders:` au niveau du tableau s'appliquent à chaque cellule au lieu du contour du tableau. Pour ne tracer que le contour, définissez `borders:` sur les cellules du bord, pour leurs côtés extérieurs uniquement.

**Les numéros de page dans les cellules de tableau commencent par « GlobalOptions ».**
Problème connu de `TableCell#page_numbers`. Placez plutôt les numéros de page dans un pied de page ou un en-tête.

**Une image JPEG ou GIF s'affiche mal ou pas du tout.**
Chaque image est stockée comme une partie PNG avec un type de contenu PNG, quel que soit son format réel : certains programmes peuvent donc ne pas afficher les autres formats. Convertissez les images en PNG avant de les ajouter.

## <a id="inspecting-a-generated-file"></a>Inspecter un fichier généré

Un `.docx` est une archive zip : vous pouvez donc regarder à l'intérieur.

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Si le problème persiste, [ouvrez une issue](https://github.com/hammad417/docxi/issues) avec un court script qui le reproduit et les versions de Ruby, docxi, nokogiri et rubyzip.

## <a id="related-pages"></a>Pages associées

- [FAQ](FAQ)
- [Bonnes pratiques](Best-Practices)
- [Installation](Installation)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
