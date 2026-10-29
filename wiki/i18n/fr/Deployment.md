<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# Utiliser docxi en production et publier des versions

docxi est une bibliothèque, pas un service : il n'y a donc rien à déployer en tant que tel. Cette page couvre ce qui compte quand vous livrez une application qui utilise docxi, et la façon dont les nouvelles versions de docxi sont publiées.

## <a id="shipping-an-app-that-uses-docxi"></a>Livrer une application qui utilise docxi

- **Dépendances.** `bundle install` installe nokogiri et rubyzip. nokogiri fournit des gems précompilées pour les plateformes Linux, macOS et Windows courantes. Sur d'autres plateformes, par exemple certaines images de conteneur minimales, elle se compile depuis les sources et nécessite un compilateur C. Voir le [guide d'installation de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Aucun logiciel de bureautique.** Les serveurs et les conteneurs n'ont besoin ni de Microsoft Word ni de LibreOffice.
- **Mémoire.** `render` construit tout le document en mémoire et le renvoie sous forme de `StringIO`. La consommation de mémoire augmente avec le nombre de pages et la taille des images.
- **Réseau et fichiers.** docxi ne lit le système de fichiers que pour les chemins d'images que vous passez à `add_media`, et n'utilise le réseau que pour les URL d'images que vous passez. Voir [Sécurité](Security).
- **Documents longs.** Générez-les dans une tâche en arrière-plan plutôt que pendant une requête web. Voir [Intégrations](Integrations).

## <a id="how-docxi-releases-are-published"></a>Comment les versions de docxi sont publiées

Ces étapes concernent la maintenance du projet. Les versions sont publiées sur RubyGems via GitHub Actions avec la publication de confiance (trusted publishing) : aucune clé d'API n'est stockée dans le dépôt.

1. Mettez à jour `lib/docxi/version.rb` et `CHANGELOG.md`, puis fusionnez dans `master`.
2. Ajoutez au commit un tag portant le numéro de version précédé de `v`, par exemple `v0.1.1`, puis poussez le tag.
3. `.github/workflows/release.yml` construit la gem et la publie sur RubyGems.

Les préversions suivent les mêmes étapes avec une version comme `0.2.0.rc1`. Bundler n'installe les préversions que si elles sont demandées explicitement. La CI (`.github/workflows/ci.yml`) s'exécute à chaque push sur `master` et pour chaque pull request.

## <a id="related-pages"></a>Pages associées

- [Installation](Installation)
- [Environnement de développement](Development-Setup)
- [Journal des modifications](Changelog)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
