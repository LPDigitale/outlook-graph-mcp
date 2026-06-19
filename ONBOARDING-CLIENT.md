# Onboarding d'un client — Outlook MCP (image mutualisée)

Procédure pour déployer le connecteur **chez un client**, dans **son** Azure, à partir de
l'**image publique mutualisée** `ghcr.io/lpdigitale/outlook-graph-mcp` — **sans ACR**, ~0–1 €/mois.

> **Garantie de traitement** : les e-mails restent dans Microsoft 365, le **jeton** d'accès reste
> dans **le Key Vault du client**. UNIY accompagne mais **ne détient ni la donnée ni la clé**.

---

## Prérequis (côté client)
- Un **abonnement Azure** actif.
- Le **compte Microsoft 365** à outiller + un **admin** pour le consentement.
- `azd` installé (ou UNIY pilote la session, écran partagé).

---

## 1. Inscrire l'app Azure AD (tenant du client)
Dans le portail Azure **du client** → **Entra ID** → **Inscriptions d'applications** → **Nouvelle** :
- Nom : `Outlook MCP`, **Locataire unique** (le sien), pas d'URI de redirection
- **Authentification** → **Autoriser les flux de client public** = **Oui**
- **Autorisations d'API** → **Microsoft Graph** → **déléguées** : `Mail.ReadWrite`, `Mail.Send`,
  `MailboxSettings.ReadWrite` → *Accorder le consentement administrateur* si le tenant l'exige
- Noter **Client ID** + **Tenant ID**

## 2. Générer une clé d'API (propre à ce client)
```powershell
[guid]::NewGuid().ToString('N') + [guid]::NewGuid().ToString('N')
```

## 3. Déployer dans l'Azure du client (image mutualisée, sans build, sans ACR)
```powershell
cd D:\Audits\outlook-graph-mcp
azd auth login                                  # se connecter à l'Azure DU CLIENT
azd env new <client>
azd env set AZURE_LOCATION       francecentral
azd env set OUTLOOK_MCP_CLIENT_ID  <client-app-id>
azd env set OUTLOOK_MCP_TENANT_ID  <client-tenant-id>
azd env set OUTLOOK_MCP_API_KEY    <clé générée à l'étape 2>
azd env set OUTLOOK_MCP_IMAGE      ghcr.io/lpdigitale/outlook-graph-mcp:0.1.0
azd provision                                   # PAS « azd up » : aucun build, tire l'image publique
```
> ⚠️ Bien `azd provision` (et non `azd up`) : en mode image mutualisée il n'y a **rien à construire**.

En sortie, noter :
- `MCP_ENDPOINT` → l'URL `/mcp` à brancher dans Claude
- `OUTLOOK_MCP_KEYVAULT_URL` → pour l'amorçage (étape 4)

## 4. Amorcer le jeton (avec le compte du client)
```powershell
$env:OUTLOOK_MCP_CLIENT_ID     = "<client-app-id>"
$env:OUTLOOK_MCP_TENANT_ID     = "<client-tenant-id>"
$env:OUTLOOK_MCP_CACHE_BACKEND = "keyvault"
$env:OUTLOOK_MCP_KEYVAULT_URL  = "<KV url renvoyée à l'étape 3>"
uv run --directory D:\Audits\outlook-graph-mcp outlook-graph-mcp login
```
→ Le **client** complète le **device-code** avec **son** compte M365 → son refresh token est écrit
dans **son** Key Vault. *(Celui qui lance `azd provision` reçoit l'accès au Key Vault ; en
accompagnement, UNIY peut lancer la commande et le client valide la connexion à l'écran.)*

## 5. Brancher dans le Claude du client
- **URL** = `MCP_ENDPOINT`, **en-tête** `x-api-key: <clé>` (ou `?key=` / `Bearer`).
- Pour la clé hors du fichier de config : variable d'env + `headersHelper` (voir `README.md`).

---

## Exploitation
| Action | Commande |
|---|---|
| Mettre à jour le connecteur | UNIY repush l'image (nouveau tag) → chez le client : `azd env set OUTLOOK_MCP_IMAGE …:<tag>` puis `azd provision` |
| Voir l'état | portail → Container App → `Log stream` ; ou `/healthz` |
| Révoquer | le client retire le consentement de l'app (le jeton devient inutile) |
| Tout supprimer | `azd down --purge` |

## Coût (chez le client)
Container App *scale-to-zero* ~0 € · Key Vault ~0 € · logs 0–1 € · **registre 0 €** (mutualisé) → **~0–1 €/mois**.

## Versionnage
Déployer les clients sur une **version figée** (`:0.1.0`), pas `:latest`, pour des montées de version maîtrisées.
