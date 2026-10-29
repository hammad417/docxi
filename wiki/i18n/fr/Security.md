<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Sécurité de docxi : signaler un problème et traiter les entrées non fiables

Cette page explique comment signaler un problème de sécurité, et quels points de sécurité connaître quand votre application transmet des saisies d'utilisateurs à docxi.

## <a id="reporting-a-vulnerability"></a>Signaler une vulnérabilité

Merci de ne pas signaler les problèmes de sécurité dans des issues publiques. Écrivez plutôt aux adresses de maintenance indiquées dans [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Indiquez la version de docxi et un moyen court de reproduire le problème.

## <a id="upgrade-from-00x"></a>Mettez à niveau depuis 0.0.x

docxi 0.0.x transmet les chaînes de chemin d'image à `Kernel#open`. `Kernel#open` exécute comme commande shell toute chaîne qui commence par `|` : quiconque contrôle un chemin d'image pourrait donc exécuter des commandes sur votre serveur. La 0.1.0 ouvre les fichiers avec `File.open` et télécharge les URL avec `URI#open`. Aucun des deux n'exécute de commande.

## <a id="untrusted-input"></a>Entrées non fiables

**Le texte peut être transmis sans risque.** Le texte, le contenu des tableaux et les autres chaînes sont écrits comme texte XML ou valeurs d'attribut, que nokogiri échappe. Les saisies des utilisateurs ne peuvent pas ajouter d'éléments au document.

**Les sources d'image demandent de la prudence.** `add_media` accepte :

| Source | Ce que fait docxi |
|---|---|
| Un IO que vous avez ouvert | Le lit |
| Un chemin de fichier | Lit ce fichier sur le disque de votre serveur |
| Une URL `http`, `https` ou `ftp` | La télécharge depuis votre serveur |

Si un utilisateur peut choisir le chemin, il peut faire intégrer au document n'importe quel fichier lisible de votre serveur, comme une configuration ou des clés. S'il peut choisir l'URL, votre serveur la demandera, ce qui peut atteindre des adresses internes (falsification de requête côté serveur, SSRF). Passez un IO que vous avez ouvert vous-même, ou n'autorisez que des chemins et des hôtes connus.

**Les liens hypertexte** pointent là où l'URL l'indique. Vérifiez les URL fournies par les utilisateurs avant de les passer à `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Comment les versions sont protégées

- Les versions sont publiées depuis GitHub Actions avec la publication de confiance (trusted publishing) de RubyGems : aucune clé d'API n'est stockée dans le dépôt.
- Les métadonnées de la gem définissent `rubygems_mfa_required` : les publications faites avec une clé d'API personnelle exigent donc une authentification multifacteur.

## <a id="related-pages"></a>Pages associées

- [Bonnes pratiques](Best-Practices)
- [Déploiement](Deployment)
- [Contribuer](Contributing)

---

[← Accueil du wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
