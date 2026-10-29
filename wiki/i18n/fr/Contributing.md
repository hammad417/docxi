<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Contribuer à docxi

Les signalements de bugs, les correctifs, la documentation et les traductions sont les bienvenus. Cette page explique comment signaler un problème et ce dont une pull request a besoin. Pour préparer une copie de travail, voir [Environnement de développement](Development-Setup).

## <a id="reporting-a-bug"></a>Signaler un bug

Ouvrez une issue sur [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) avec :

- un court script Ruby qui reproduit le problème
- les versions de Ruby, docxi, nokogiri et rubyzip (et de Rails, si vous l'utilisez)
- ce que vous attendiez, et ce que vous avez obtenu. Pour un document défectueux, précisez avec quel programme vous l'avez ouvert

Signalez plutôt les problèmes de sécurité en privé. Voir [Sécurité](Security).

## <a id="pull-requests"></a>Pull requests

1. Forkez le dépôt et créez une branche.
2. Faites la modification en respectant le style du code environnant.
3. Ajoutez ou mettez à jour des tests, puis lancez `bundle exec rake test`.
4. Si le XML généré change volontairement, lancez `bundle exec rake golden` et expliquez le changement dans la pull request. Les fichiers de référence protègent la sortie dont dépendent les utilisateurs actuels.
5. Mettez à jour `CHANGELOG.md` et la documentation si le comportement change.
6. Ouvrez la pull request. La CI exécute les tests sous Ruby 2.5 à 4.0 et Rails 5.2 à 8.1, et vérifie que LibreOffice ouvre les documents d'exemple.

Les changements qui modifient des documents existants, comme les correctifs prévus pour la 0.2.0, doivent faire l'objet de leur propre pull request, afin que le diff des fichiers de référence reste facile à relire.

## <a id="documentation-and-translations"></a>Documentation et traductions

- Le README anglais est la source de 15 traductions dans `docs/i18n/`. Suivez `docs/i18n/TRANSLATION_GUIDE.md` dans le dépôt quand vous le modifiez.
- Ce wiki est rédigé dans le dossier `wiki/` du dépôt et publié à partir de là. `wiki/README.md` explique comment.

## <a id="related-pages"></a>Pages associées

- [Environnement de développement](Development-Setup)
- [Architecture](Architecture)
- [Feuille de route](Roadmap)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
