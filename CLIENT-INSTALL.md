# Installation — Connecteur Outlook MCP

Déployez ce connecteur **dans votre propre Azure**. Il permet à un assistant IA (Claude)
de gérer votre boîte Outlook (dossiers, catégories, règles, messages, réponses automatiques)
via Microsoft Graph.

> 🔒 **Vos données restent chez vous.** Les e-mails ne quittent pas Microsoft 365 ; le jeton
> d'accès est stocké chiffré dans **votre** Key Vault. L'éditeur (UNIY) n'héberge rien et n'a
> accès ni à vos e-mails ni à votre clé.

---

## Prérequis
- Un **abonnement Azure** (droit de créer des ressources).
- Un compte **Microsoft 365** (la boîte à outiller) + un **administrateur** pour le consentement.
- Trois outils gratuits :
  - [git](https://git-scm.com/downloads)
  - [Azure Developer CLI (`azd`)](https://aka.ms/install-azd)
  - [uv (Python)](https://docs.astral.sh/uv/getting-started/installation/) — utilisé une fois pour la connexion

---

## 1. Récupérer le projet
```bash
git clone https://github.com/LPDigitale/outlook-graph-mcp.git
cd outlook-graph-mcp
```

## 2. Inscrire l'application (votre Microsoft Entra ID)
Dans le [portail Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Inscriptions d'applications** → **Nouvelle inscription** :
- Nom : `Outlook MCP` · **Comptes de cet annuaire uniquement** · pas d'URI de redirection → **S'inscrire**
- **Authentification** → **Autoriser les flux de client public** = **Oui** → Enregistrer
- **Autorisations d'API** → **Microsoft Graph** → **Autorisations déléguées** : `Mail.ReadWrite`,
  `Mail.Send`, `MailboxSettings.ReadWrite` → **Accorder le consentement administrateur**
- Notez l'**ID d'application (client)** et l'**ID de l'annuaire (locataire)**

## 3. Déployer dans votre Azure (≈ 3 min, ~0–1 €/mois)
```powershell
azd auth login
azd env new outlook-mcp
azd env set AZURE_LOCATION       francecentral
azd env set OUTLOOK_MCP_CLIENT_ID  <ID d'application>
azd env set OUTLOOK_MCP_TENANT_ID  <ID de locataire>
azd env set OUTLOOK_MCP_API_KEY    <une clé secrète, voir ci-dessous>
azd env set OUTLOOK_MCP_IMAGE      ghcr.io/lpdigitale/outlook-graph-mcp@sha256:f93ecccc6ac389f03b0b423516cc53664538656d887bec2496ec9954351621dc
azd provision
```
**Générer une clé d'API** (à coller dans `OUTLOOK_MCP_API_KEY`) :
```powershell
[guid]::NewGuid().ToString('N') + [guid]::NewGuid().ToString('N')
```
À la fin, notez les **`Outputs`** : `MCP_ENDPOINT` et `OUTLOOK_MCP_KEYVAULT_URL`.

> Le déploiement tire une **image publique** (aucune construction) et ne crée **aucun registre** → coût minimal.
> L'image est **épinglée par empreinte `@sha256:…` (immuable)** : vous exécutez exactement la version
> publiée et revue (ici la `v0.1.0`), insensible à toute modification ultérieure d'un tag.

## 4. Connexion (une seule fois)
Connecte **votre** compte Microsoft et dépose le jeton dans **votre** Key Vault :
```powershell
$env:OUTLOOK_MCP_CLIENT_ID     = "<ID d'application>"
$env:OUTLOOK_MCP_TENANT_ID     = "<ID de locataire>"
$env:OUTLOOK_MCP_CACHE_BACKEND = "keyvault"
$env:OUTLOOK_MCP_KEYVAULT_URL  = "<OUTLOOK_MCP_KEYVAULT_URL de l'étape 3>"
uv run --directory . outlook-graph-mcp login
```
Suivez le **code** affiché sur `https://microsoft.com/devicelogin`, connectez-vous avec **votre**
compte → c'est terminé. *(En cas d'erreur 403 sur le Key Vault juste après le déploiement,
attendez ~1 min puis réessayez.)*

## 5. Brancher dans Claude
Ajoutez un **serveur MCP distant (HTTP)** :
- **URL** : la valeur `MCP_ENDPOINT`
- **En-tête d'authentification** : `x-api-key: <votre clé>`
  *(également accepté : `?key=<clé>` dans l'URL, ou `Authorization: Bearer <clé>`)*

Vérifiez : demandez à Claude « liste mes dossiers Outlook ».

---

## Sécurité & exploitation
- Permissions **déléguées** (`/me` uniquement) : le connecteur agit **en votre nom**, sur **votre** boîte.
- L'endpoint est protégé par votre **clé d'API** (à garder secrète).
- **Révoquer** : retirez le consentement de l'app dans Entra ID → le jeton devient inutile.
- **Mettre à jour** : UNIY publie une nouvelle version (nouveau digest) → `azd env set OUTLOOK_MCP_IMAGE ghcr.io/lpdigitale/outlook-graph-mcp@sha256:<nouveau digest>` puis `azd provision`.
- **Tout supprimer** : `azd down --purge`.

## Besoin d'aide ?
UNIY peut réaliser l'installation **avec vous** (accompagnement) — notamment l'étape 4.
