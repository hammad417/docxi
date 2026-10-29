<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Environnement de développement : lancer les tests et les outils de docxi

Cette page vous aide à faire fonctionner une copie de travail de docxi, et explique les tests, les documents d'exemple et la configuration de la CI.

## <a id="get-the-code"></a>Récupérer le code

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

N'importe quel Ruby à partir de la 2.5 convient pour le développement. Les dépendances de développement sont minitest et rake.

## <a id="rake-tasks"></a>Tâches Rake

| Commande | Ce qu'elle fait |
|---|---|
| `bundle exec rake test` | Lance la suite de tests (la tâche par défaut) |
| `bundle exec rake samples` | Écrit les documents d'exemple dans `tmp/samples/*.docx`, pour les ouvrir dans Word, LibreOffice ou Google Docs |
| `bundle exec rake golden` | Réécrit `test/golden` à partir du code actuel. À utiliser uniquement après un changement volontaire de la sortie |
| `bundle exec rake build` | Construit la gem dans `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Lancer les tests avec Rails chargé

Le dossier `gemfiles/` contient un Gemfile par version de Rails, ainsi que quelques-uns pour d'anciennes versions de rubyzip et de nokogiri :

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` charge tout le framework avant les tests. Il compare aussi les règles de `blank?` de docxi avec celles d'ActiveSupport.

## <a id="what-the-tests-cover"></a>Ce que couvrent les tests

| Fichier | Vérifie |
|---|---|
| `test/test_golden_output.rb` | Chaque partie de deux documents d'exemple correspond à ce que la 0.0.6 générait |
| `test/test_package.rb` | Le XML est bien formé, les relations pointent vers des parties existantes, chaque partie a un type de contenu et il n'y a pas d'enregistrements Zip64 |
| `test/test_media.rb` | Les images depuis des chemins, des objets IO et des URL, et le fait qu'une chaîne `\|` n'est jamais exécutée comme commande |
| `test/test_fixed_features.rb` | `tab`, `frame` et les numéros de page dans les en-têtes |
| `test/test_util.rb` | `Docxi::Util.blank?` suit les règles d'ActiveSupport |
| `test/test_rails_independence.rb` | docxi fonctionne sans ActiveSupport |
| `test/test_warnings.rb` | La génération n'émet aucun avertissement de dépréciation depuis `lib/` |

Les documents d'exemple se trouvent dans `test/support/sample_documents.rb`, et `test/support/package.rb` lit un `.docx` directement à partir de ses structures zip.

## <a id="continuous-integration"></a>Intégration continue

`.github/workflows/ci.yml` s'exécute à chaque push sur `master` et pour chaque pull request. Elle couvre :

- Ruby 2.5, 2.6 et 4.0 sans Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) et 8.1 (4.0)
- Rails `main`, dont l'échec est toléré
- rubyzip 1.3 et 2.4, ainsi que nokogiri 1.11
- un job qui convertit les documents d'exemple en PDF avec LibreOffice

Les publications de versions sont décrites dans [Déploiement](Deployment).

## <a id="related-pages"></a>Pages associées

- [Contribuer](Contributing)
- [Architecture](Architecture)
- [Déploiement](Deployment)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
