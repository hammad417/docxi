<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Installer docxi : prérequis, configuration de Bundler et mise à niveau

docxi est une gem écrite entièrement en Ruby. Son installation n'ajoute que deux dépendances, nokogiri et rubyzip, et rien d'autre. Aucun logiciel de bureautique n'est nécessaire.

## <a id="requirements"></a>Prérequis

| Prérequis | Version |
|---|---|
| Ruby | 2.5 ou plus récent |
| nokogiri | `>= 1.8, < 2`, installé automatiquement |
| rubyzip | `>= 1.3, < 4`, installé automatiquement |
| Rails | Non requis |

## <a id="install-with-bundler"></a>Installer avec Bundler

Ajoutez la gem à votre Gemfile :

```ruby
gem "docxi", "~> 0.1"
```

Puis installez-la :

```sh
bundle install
```

## <a id="install-without-bundler"></a>Installer sans Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Vérifier l'installation

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Dans un projet Bundler, lancez plutôt `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. La commande affiche la version installée, par exemple `0.1.0`.

## <a id="upgrading-from-00x"></a>Mettre à niveau depuis 0.0.x

Si votre Gemfile fixe docxi sur une 0.0.x, par exemple avec `"~> 0.0.6"`, remplacez la ligne par `gem "docxi", "~> 0.1"`. Lancez ensuite `bundle update docxi`. Aucune modification du code n'est nécessaire, et la 0.1.0 génère les mêmes documents que la 0.0.6 : la suite de tests compare la sortie à des fichiers générés par la 0.0.6.

Ce qui a changé en interne :

- nokogiri et rubyzip peuvent désormais être dans n'importe quelle version des plages ci-dessus, si bien que docxi s'installe à côté des versions actuelles de Rails.
- docxi n'a plus besoin qu'ActiveSupport soit chargé.
- Les archives sont écrites sans enregistrements Zip64.
- Les images sont chargées sans `Kernel#open`, donc les URL fonctionnent sous Ruby 3.
- `tab`, `frame` et les numéros de page dans les en-têtes fonctionnent.

Les applications qui tournent sur une version de Ruby antérieure à 2.5 continuent de recevoir la 0.0.6, car la 0.1.0 déclare Ruby 2.5 comme version minimale. Le [journal des modifications](Changelog) liste toutes les modifications.

## <a id="versions-tested-in-ci"></a>Versions testées en CI

Chaque version de Rails est testée avec le framework complet chargé à côté de docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

La CI exécute aussi :

- les tests sans Rails sous Ruby 2.5, 2.6 et 4.0
- les tests avec rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) et nokogiri 1.11 (Ruby 2.5)
- une vérification que LibreOffice peut convertir les documents d'exemple

## <a id="platform-notes"></a>Remarques sur les plateformes

- nokogiri fournit des gems précompilées pour les plateformes Linux, macOS et Windows courantes. Quand aucune ne correspond, elle se compile depuis les sources et nécessite un compilateur C. Voir le [guide d'installation de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- La CI tourne sur Ubuntu 24.04, et les tests ont aussi été exécutés sous macOS. Windows n'est pas testé.

## <a id="if-installation-fails"></a>Si l'installation échoue

Un conflit de Bundler sur nokogiri ou rubyzip signifie presque toujours qu'une ancienne docxi 0.0.6 est en cours de résolution. Voir [Dépannage](Troubleshooting#installation).

## <a id="related-pages"></a>Pages associées

- [Premiers pas](Getting-Started)
- [Configuration](Configuration)
- [Déploiement](Deployment)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
