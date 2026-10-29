<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Journal des modifications et historique des versions de docxi

Cette page résume chaque version. Les notes complètes se trouvent dans [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) dans le dépôt, qui fait référence en cas de divergence.

## <a id="releases"></a>Versions

| Version | Date | Notes |
|---|---|---|
| 0.1.1 | 2026-10-29 | Résumé et description du gem plus justes, et lien vers la documentation. Aucun changement de code |
| 0.1.0 | 2026-10-28 | Fonctionne avec les versions actuelles de Ruby, Rails et rubyzip. Génère les mêmes documents que la 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Préversion de la 0.1.0 |
| 0.0.6 | 2016-10-18 | Dernière version avant la mise à jour 0.1.0 |
| 0.0.2 | 2016-10-14 | Aucune note de version n'a été enregistrée |
| 0.0.1 | 2016-10-14 | Première version. Aucune note de version n'a été enregistrée |

Les versions 0.0.3 à 0.0.5 n'ont jamais été publiées sur RubyGems.

## <a id="011"></a>0.1.1

Aucun changement de code, et les documents générés sont identiques à ceux de la 0.1.0.

- Le résumé et la description du gem disent maintenant ce que fait docxi. Ils le présentaient comme un générateur de modèles.
- RubyGems renvoie vers ce wiki comme documentation du gem.

## <a id="010"></a>0.1.0

**Compatibilité**

- Nécessite Ruby 2.5 ou plus récent. Les applications sur des versions plus anciennes de Ruby restent sur la 0.0.6.
- Fonctionne avec ou sans Rails, et est testée avec Rails 5.2 à 8.1.
- nokogiri `>= 1.8, < 2` au lieu de `>= 1.6.3, < 1.7`, qui entrait en conflit avec toutes les versions de Rails depuis la 6.0.
- rubyzip `>= 1.3, < 4` au lieu de `~> 0.9.9`.
- N'appelle plus `blank?` et `present?` d'ActiveSupport, qui ne fonctionnaient que si Rails les avait chargés.

**Correctifs**

- Les documents écrits avec rubyzip 3 ne contiennent plus d'enregistrements Zip64, que Google Docs et les anciennes versions de LibreOffice rejettent.
- `add_media` avec une URL d'image fonctionne de nouveau sous Ruby 3.
- `add_media` n'exécute plus comme commande shell une chaîne qui commence par `|`.
- `tab` et `frame` ne lèvent plus `NameError`.
- `document.frame("text")` crée une zone de texte contenant le texte.
- `page_numbers` dans un en-tête ne lève plus `NameError`.

**Gem**

- Changement de maintenance, et nouveaux liens vers la page d'accueil et le journal des modifications.
- Les publications exigent la MFA.

## <a id="related-pages"></a>Pages associées

- [Feuille de route](Roadmap)
- [Installation](Installation#upgrading-from-00x)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
