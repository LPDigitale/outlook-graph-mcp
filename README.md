# Outlook Graph MCP

Serveur **MCP** qui pilote votre boîte **Outlook** via **Microsoft Graph**, en écriture, avec des
**permissions déléguées** (tout se passe sur votre propre boîte `/me`). Il expose chaque action du
tableau ci-dessous comme un **outil** appelable par Claude.

> Contrairement au connecteur Outlook « lecture seule » (recherche d'e-mails), celui-ci **crée,
> modifie et envoie** : dossiers, catégories, règles de tri, messages, réponses automatiques, etc.

---

## Couverture (votre tableau → outils MCP)

| Domaine | Action | Outil MCP | Endpoint Graph |
|---|---|---|---|
| 📁 Dossiers | Lister | `list_mail_folders`, `list_child_folders` | `GET /me/mailFolders` |
| | Créer (racine / sous-dossier) | `create_mail_folder` | `POST /me/mailFolders[/{id}/childFolders]` |
| | Renommer | `rename_mail_folder` | `PATCH /me/mailFolders/{id}` |
| | Supprimer | `delete_mail_folder` ⚠️ | `DELETE /me/mailFolders/{id}` |
| | Voir le contenu | `list_folder_messages` | `GET /me/mailFolders/{id}/messages` |
| 🏷️ Catégories | Lister | `list_categories` | `GET /me/outlook/masterCategories` |
| | Créer (+ couleur) | `create_category` | `POST /me/outlook/masterCategories` |
| | Modifier | `update_category` | `PATCH …/masterCategories/{id}` |
| | Appliquer à un message | `update_message` (param `categories`) | `PATCH /me/messages/{id}` |
| ⚙️ Règles | Lister | `list_inbox_rules` | `GET /me/mailFolders/inbox/messageRules` |
| | Créer | `create_inbox_rule` | `POST …/messageRules` |
| | Modifier / activer / désactiver | `update_inbox_rule` | `PATCH …/messageRules/{id}` |
| | Supprimer | `delete_inbox_rule` | `DELETE …/messageRules/{id}` |
| ✉️ Messages | Rechercher / lister | `list_messages`, `list_folder_messages` | `GET /me/messages` |
| | Lire (+ pièces jointes) | `get_message` | `GET /me/messages/{id}` |
| | Lu/non lu, importance, classement, drapeau | `update_message` | `PATCH /me/messages/{id}` |
| | Déplacer | `move_message` | `POST /me/messages/{id}/move` |
| | Supprimer | `delete_message` | `DELETE /me/messages/{id}` |
| | Envoyer (+ pièces jointes) | `send_mail` | `POST /me/sendMail` |
| | Répondre / Répondre à tous (envoi immédiat) | `reply_message` | `POST …/reply` · `…/replyAll` |
| | Répondre EN BROUILLON, dans le fil (non envoyé) | `create_reply_draft` | `POST …/createReply` · `…/createReplyAll` |
| | Transférer (envoi immédiat) | `forward_message` | `POST …/forward` |
| | Transférer EN BROUILLON (non envoyé) | `create_forward_draft` | `POST …/createForward` |
| | Brouillon nouveau fil (créer / modifier / envoyer) | `create_draft`, `update_draft`, `send_draft` | `POST`/`PATCH /me/messages` |
| | Ajouter une pièce jointe | `add_attachment` | `POST …/attachments` |
| 🔧 Paramètres | Réponses automatiques / absence | `set_automatic_replies` | `PATCH /me/mailboxSettings` |
| | Fuseau, langue, formats | `update_mailbox_settings`, `get_mailbox_settings` | `GET`/`PATCH /me/mailboxSettings` |
| 🩺 Diagnostic | Tester la connexion | `whoami` | `GET /me` |

> **Bonus** : `delete_mail_folder` lève la limite « suppression de dossier impossible » mentionnée
> précédemment. C'est une action **destructive** — Claude vous demandera confirmation.

---

## Prérequis

- Windows + **Python ≥ 3.10** et **[uv](https://docs.astral.sh/uv/)** (déjà installés sur votre poste).
- Un compte **Microsoft 365 professionnel** (ex. `…@relationdigitale.com`).
- Les droits de créer une **inscription d'application** dans Azure (ou un admin qui consent — voir étape A).

---

## Étape A — Inscrire l'application Azure AD (≈ 10 min, une seule fois)

1. Ouvrez le **[portail Azure](https://portal.azure.com)** → **Microsoft Entra ID** → **Inscriptions
   d'applications** → **Nouvelle inscription**.
2. **Nom** : `Outlook MCP` (libre). **Types de comptes pris en charge** :
   *Comptes dans cet annuaire d'organisation uniquement* (mono-tenant). **URI de redirection** : laissez **vide**.
   → **S'inscrire**.
3. Sur la page **Vue d'ensemble**, copiez :
   - **ID d'application (client)** → `OUTLOOK_MCP_CLIENT_ID`
   - **ID de l'annuaire (locataire)** → `OUTLOOK_MCP_TENANT_ID`
4. **Authentification** → **Paramètres avancés** → **Autoriser les flux de client public** → **Oui** → **Enregistrer**.
   *(Indispensable pour la connexion device-code.)*
5. **Autorisations d'API** → **Ajouter une autorisation** → **Microsoft Graph** → **Autorisations déléguées**,
   puis ajoutez :
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `MailboxSettings.ReadWrite`
   *(`User.Read` est déjà présent.)*
6. Si votre locataire l'exige, cliquez **Accorder un consentement administrateur pour …**.
   Sinon, le consentement se fera à votre première connexion.

> 💡 Vous pouvez aussi me demander de **vous guider en direct dans le navigateur** pour cette étape.

---

## Étape B — Installer le serveur

```powershell
cd D:\Audits\outlook-graph-mcp
uv sync
```

`uv` crée un environnement isolé et installe les dépendances (`mcp`, `msal`, `msal-extensions`, `httpx`).

---

## Étape C — Se connecter (une fois)

Renseignez vos identifiants Azure puis lancez la connexion **device-code** :

```powershell
$env:OUTLOOK_MCP_CLIENT_ID = "VOTRE_CLIENT_ID"
$env:OUTLOOK_MCP_TENANT_ID = "VOTRE_TENANT_ID"
uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp login
```

Le terminal affiche un message du type :
*« To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code XXXX-YYYY. »*
→ ouvrez la page, saisissez le code, validez avec votre compte Microsoft 365. À la fin :

```
Connecté : Ludovic Perrot <ludovic.perrot@relationdigitale.com>
```

Le **jeton est mis en cache et chiffré** (DPAPI Windows) dans `%USERPROFILE%\.outlook-graph-mcp`.
Le serveur le rafraîchit ensuite tout seul — vous n'avez plus à vous reconnecter.

*(Variante navigateur : `… outlook-graph-mcp login --interactive`. Elle nécessite d'ajouter, à l'étape A,
un URI de redirection `http://localhost` sous « Mobile et applications de bureau ».)*

---

## Étape D — Brancher dans Claude

### Claude Code (CLI)
Soit via la commande :
```powershell
claude mcp add outlook-graph --env OUTLOOK_MCP_CLIENT_ID=VOTRE_CLIENT_ID --env OUTLOOK_MCP_TENANT_ID=VOTRE_TENANT_ID -- uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp
```
Soit en copiant `.mcp.json.example` vers `.mcp.json` (ou dans votre config utilisateur) et en y mettant vos IDs.

### Claude Desktop
Éditez `claude_desktop_config.json`
(`%APPDATA%\Claude\claude_desktop_config.json`) et ajoutez le bloc de `.mcp.json.example`
(la clé racine est `mcpServers`). Redémarrez Claude Desktop.

Une fois branché, demandez par exemple :
> « Liste mes dossiers Outlook », « Crée une règle qui classe les e-mails de @client.com dans le dossier *test* »,
> « Active mon absence du bureau du 1er au 15 août ».

---

## 🚀 Déploiement sur Azure (serveur MCP distant)

Héberge le serveur sur **Azure Container Apps** (scale-to-zero), derrière une URL HTTPS publique
protégée par clé d'API. Le jeton délégué est conservé chiffré dans **Azure Key Vault** ; l'app y
accède via une **identité managée**. Modèle d'auth Graph **inchangé** (`/me`, délégué) — vous restez
le seul propriétaire de l'accès, **sans consentement admin**.

Tout est décrit par `azure.yaml` + `infra/*.bicep` et déployé par `azd`.

### Prérequis
- `azd` installé (Azure Developer CLI) ✅
- Une **app Azure AD** (étape A ci-dessus) → `Client ID` + `Tenant ID`
- Un **abonnement Azure** où vous pouvez créer des ressources

### 1. Connecter azd à votre abonnement
```powershell
azd auth login
```

### 2. Configurer l'environnement
```powershell
cd D:\Audits\outlook-graph-mcp
azd env new outlook-mcp
azd env set AZURE_LOCATION       francecentral
azd env set OUTLOOK_MCP_CLIENT_ID  VOTRE_CLIENT_ID
azd env set OUTLOOK_MCP_TENANT_ID  VOTRE_TENANT_ID
azd env set OUTLOOK_MCP_API_KEY    VOTRE_CLE_API
```

### 3. Provisionner + déployer
```powershell
azd up
```
Crée le groupe de ressources, le registre, l'environnement Container Apps, le Key Vault et l'app,
puis construit l'image (dans le cloud) et la déploie. En sortie, notez :
- `MCP_ENDPOINT` → l'URL à brancher dans Claude (`https://….azurecontainerapps.io/mcp`)
- `OUTLOOK_MCP_KEYVAULT_URL` → pour l'amorçage du jeton (étape 4)

### 4. Amorcer le jeton (une seule fois)
Connecte **votre** compte M365 et écrit le refresh token dans le Key Vault (vous y avez le rôle
*Secrets Officer*, accordé par le déploiement) :
```powershell
$env:OUTLOOK_MCP_CLIENT_ID   = "VOTRE_CLIENT_ID"
$env:OUTLOOK_MCP_TENANT_ID   = "VOTRE_TENANT_ID"
$env:OUTLOOK_MCP_CACHE_BACKEND = "keyvault"
$env:OUTLOOK_MCP_KEYVAULT_URL  = "https://kv-….vault.azure.net/"
uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp login
```
*(En cas d'erreur 403 sur le Key Vault juste après `azd up`, attendez ~1 min — propagation RBAC — puis réessayez.)*

### 5. Brancher dans Claude (connecteur MCP distant)
Ajoutez un **serveur MCP HTTP** pointant sur `MCP_ENDPOINT`, avec l'en-tête d'authentification :
```
x-api-key: VOTRE_CLE_API
```
La clé est aussi acceptée via `?key=VOTRE_CLE_API` ou `Authorization: Bearer VOTRE_CLE_API`.

### Exploitation
| Action | Commande |
|---|---|
| Mettre à jour le code déployé | `azd deploy` |
| Voir les logs | `azd monitor` ou portail → Container App → Log stream |
| Tout supprimer | `azd down --purge` |

**Coûts** : Container Apps *scale-to-zero* (≈ 0 € au repos), + Log Analytics / ACR Basic / Key Vault =
quelques € / mois. **Cold start** : la 1ʳᵉ requête après inactivité prend quelques secondes (l'app se réveille).

**Durcissement (optionnel)** : remplacer la clé d'API par **EasyAuth / Entra ID** (App Service
Authentication sur Container Apps) ou **API Management**, et isoler le réseau (VNet) si besoin.

---

## Permissions & portée

- Permissions **déléguées** uniquement : `Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite` (+ `User.Read`).
- Toutes les actions ciblent **votre** boîte (`/me`). Aucune action sur d'autres boîtes.
- Aucun secret client n'est stocké (client **public** + device-code). Seul un *refresh token* chiffré est conservé localement.

## Sécurité

- Cache de jeton chiffré au repos via **msal-extensions** (DPAPI sous Windows).
- `uv run … logout` efface le cache. Révocation côté Azure : *Mes connexions* / portail Entra.

## Limites connues

- Pièces jointes **≤ 3 Mo** par appel direct (au-delà, une *upload session* serait nécessaire — non implémentée).
- `$search` (recherche plein texte) ne se combine pas avec `$filter`/`$orderby` (limite Graph).
- Pas d'accès aux boîtes partagées/d'autres utilisateurs (périmètre `/me` volontaire).

## Dépannage

| Symptôme | Cause / solution |
|---|---|
| `OUTLOOK_MCP_CLIENT_ID manquant` | Variables d'env non transmises — vérifiez le bloc `env` de la config MCP. |
| `Aucun jeton valide en cache` | Lancez `… outlook-graph-mcp login`. |
| Erreur de consentement au login | Un admin doit cliquer **Accorder un consentement administrateur** (étape A.6). |
| `Allow public client flows` | Doit être **Oui** dans Authentification (étape A.4). |
| `ErrorAccessDenied` sur une action | Permission déléguée manquante — réajoutez-la (étape A.5) puis reconnectez-vous. |

---

### Commandes utiles

```powershell
# Tester la connexion
uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp whoami
# Se déconnecter (vider le cache)
uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp logout
```
