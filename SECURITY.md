# Sécurité — Outlook Graph MCP

## Modèle de sécurité
- **Permissions déléguées, périmètre `/me`** : le connecteur agit au nom de l'utilisateur connecté,
  sur **sa seule** boîte. Scopes minimaux : `Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite`.
- **Aucune donnée stockée** : le serveur est un passe-plat vers Microsoft Graph. Les e-mails restent
  dans Microsoft 365 ; ils transitent en mémoire le temps d'une requête, **ne sont jamais persistés**
  (ni base, ni disque). Les logs ne contiennent **pas** de contenu d'e-mail (méthode + chemin seulement).
- **Seul secret au repos = le jeton délégué**, chiffré dans **Azure Key Vault**, lu par l'app via une
  **identité managée**. En déploiement *chez le client*, ce jeton ne quitte jamais son abonnement.
- **Endpoint protégé par clé d'API** (en-tête `x-api-key`). À transmettre/stocker comme un mot de passe.
- **Isolation** : chaque déploiement est indépendant (sa propre app Entra, sa clé, son jeton, son Azure).
- **Révocation** : retirer le consentement de l'app dans Entra ID invalide immédiatement l'accès.

## Chaîne d'approvisionnement (image)
- Les déploiements tirent une **image publique** `ghcr.io/lpdigitale/outlook-graph-mcp`.
- **Épinglez par empreinte** `@sha256:…` (immuable), pas seulement par tag — voir `CLIENT-INSTALL.md`.
- L'image ne contient **aucun secret** (les secrets sont injectés à l'exécution via Key Vault / variables).

## Durcissement du compte éditeur (UNIY) — check-list
- [ ] **2FA** activé sur le compte GitHub qui publie l'image (obligatoire).
- [ ] Migration vers une **organisation GitHub UNIY** (accès restreint, rôles).
- [ ] **Branch protection** sur `main` (revue obligatoire, pas de force-push).
- [ ] Publication de l'image via **token à portée minimale et durée courte** (ou GitHub Actions + OIDC).
- [ ] **Secret scanning + push protection** activés (fait sur ce repo).
- [ ] **Dependabot** (alertes + correctifs) activé (fait).
- [ ] (Cible) **Signature d'image** (cosign / sigstore) → provenance vérifiable par les clients.

## Limite connue
L'authentification de l'endpoint par **clé statique** est le maillon le plus faible : en cas de fuite
de l'URL + clé, accès à la boîte. Évolution prévue : **OAuth / Entra ID** (chaque utilisateur
s'authentifie avec son compte, plus de clé statique).

## Signaler une vulnérabilité
Contact : ouvrir un *security advisory* privé sur le dépôt, ou écrire à l'éditeur (UNIY).
Merci de **ne pas** divulguer publiquement avant correction.
