<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki docxi : générer des documents Word (.docx) en Ruby

docxi est une gem Ruby qui génère des fichiers Microsoft Word `.docx` à partir de code. Elle prend en charge les paragraphes, le texte mis en forme, les tableaux, les images, les en-têtes et pieds de page, les numéros de page, les liens hypertexte et une table des matières simple. Elle ne nécessite ni Microsoft Word, ni LibreOffice, ni modèle, et fonctionne aussi bien en Ruby pur que dans les applications Ruby on Rails.

Le [README](https://github.com/hammad417/docxi#readme) en donne une présentation courte. Ce wiki couvre toutes les options, propose des recettes prêtes à copier, explique le fonctionnement interne de docxi et vous aide quand quelque chose ne va pas.

Version actuelle : **0.1.1**, pour Ruby 2.5 ou plus récent. Consultez le [journal des modifications](Changelog).

## <a id="start-here"></a>Par où commencer

**Pour découvrir docxi**

- [Premiers pas](Getting-Started) : installer la gem et créer votre premier document
- [Installation](Installation) : prérequis, Bundler et mise à niveau depuis 0.0.x
- [Exemples d'utilisation](Usage-Examples) : rapports, factures, lettres et tableaux construits à partir de données, prêts à copier

**Construire des documents**

- [Fonctionnalités](Features) : chaque élément et les options qu'il accepte
- [Configuration](Configuration) : marges, mise en page et unités
- [Référence de l'API](API-Reference) : classes et méthodes en un coup d'œil
- [Bonnes pratiques](Best-Practices) : des habitudes qui évitent les mauvaises surprises

**Quand quelque chose ne va pas**

- [Dépannage](Troubleshooting) : messages d'erreur et documents qui ne s'affichent pas correctement
- [FAQ](FAQ) : réponses courtes aux questions fréquentes
- [Sécurité](Security) : signaler un problème et traiter les entrées non fiables

**Pour aller plus loin**

- [Architecture](Architecture) : comment des appels Ruby deviennent un paquet `.docx`
- [Intégrations](Integrations) : contrôleurs Rails, applications Rack et tâches en arrière-plan
- [Déploiement](Deployment) : utiliser docxi en production, et publication des versions
- [Contribuer](Contributing) et [Environnement de développement](Development-Setup)
- [Feuille de route](Roadmap), [Journal des modifications](Changelog) et [Crédits](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>Ce que fait docxi, et ce qu'elle ne fait pas

docxi crée de nouveaux documents. Elle ne peut ni ouvrir ni modifier un `.docx` existant, ni remplir un modèle Word, ni faire de publipostage. Elle ne propose ni listes numérotées, ni styles de titre, ni table des matières automatique, et les pages sont toujours au format US Letter en orientation portrait. Les problèmes connus prévus pour la 0.2.0 sont listés dans la [feuille de route](Roadmap).

## <a id="credits"></a>Crédits

La maintenance de docxi est assurée par [Hammad Maqbool](https://github.com/hammad417) chez [Phaedra Solutions](https://www.phaedrasolutions.com/). Le projet est né en 2016 à partir de [docxer](https://github.com/gtimti/docxer), de Timur Gabdrakipov. Voir [Crédits](Credits) pour plus de détails.

## <a id="languages"></a>Langues

La version anglaise est la version de référence de ce wiki. Les traductions sont accessibles depuis la barre latérale et en haut de chaque page.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Signaler un problème](https://github.com/hammad417/docxi/issues)
