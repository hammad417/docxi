<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# FAQ docxi : documents Word en Ruby et Rails

Des réponses courtes aux questions les plus fréquentes. Chaque réponse renvoie vers une explication plus détaillée.

**Faut-il installer Microsoft Word ou LibreOffice ?**
Non. docxi écrit lui-même le XML et l'archive zip.

**docxi nécessite-t-il Rails ?**
Non. Il fonctionne dans n'importe quel programme Ruby. Les tests s'exécutent aussi avec Rails 5.2 à 8.1 chargé. Voir [Intégrations](Integrations).

**Comment générer un document Word en Ruby on Rails ?**
Construisez-le dans un contrôleur ou une tâche en arrière-plan, puis envoyez `word.render.string` avec `send_data`. Voir [Intégrations](Integrations).

**Quelles versions de Ruby et de Rails sont prises en charge ?**
Ruby 2.5 et plus récent. Rails est facultatif, et la CI teste Rails 5.2 à 8.1. Voir [Installation](Installation).

**docxi peut-il modifier un .docx existant ou remplir un modèle Word ?**
Non, il ne fait que créer de nouveaux documents. Pour des modèles avec des champs de fusion, utilisez une gem basée sur des modèles, comme sablon.

**La table des matières est-elle générée à partir des titres ?**
Non. Vous ajoutez chaque entrée et son numéro de page. Voir [Fonctionnalités](Features#table-of-contents).

**Puis-je utiliser des listes numérotées ou des styles de titre ?**
Pas dans la 0.1.0. Les paragraphes peuvent être des éléments de liste à puces (`ul: true`), et vous pouvez mettre en forme les titres avec des options de texte comme `size:` et `bold:`.

**Puis-je changer la taille ou l'orientation de la page ?**
Pas encore. Les pages sont au format US Letter en portrait, et seules les marges sont modifiables. Voir [Configuration](Configuration).

**Quels formats d'image fonctionnent ?**
Le PNG. Chaque image est stockée comme une partie PNG.

**Puis-je ajouter une image depuis une URL ?**
Oui. `add_media` accepte les URL `http`, `https` et `ftp`, et les télécharge lorsque vous appelez `render`. Ne passez pas d'URL fournies par les utilisateurs sans les vérifier. Voir [Sécurité](Security).

**Puis-je écrire le document directement dans un fichier ?**
`render` renvoie un `StringIO`. Écrivez son `.string` avec `File.binwrite("out.docx", word.render.string)`.

**docxi peut-il générer de très gros documents en streaming ?**
Non. Tout le document est construit en mémoire lorsque vous appelez `render`. Voir [Déploiement](Deployment).

**Puis-je définir des paragraphes de droite à gauche ?**
La 0.1.0 n'a pas d'option de droite à gauche. Le texte est écrit tel que vous le passez.

**Où sont listés les problèmes connus ?**
Dans la [feuille de route](Roadmap) et dans la section « Limitations and known issues » (limites et problèmes connus) du README.

## <a id="related-pages"></a>Pages associées

- [Dépannage](Troubleshooting)
- [Fonctionnalités](Features)
- [Premiers pas](Getting-Started)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
