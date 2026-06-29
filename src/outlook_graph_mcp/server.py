"""Serveur MCP Outlook : tous les outils de configuration via Microsoft Graph.

Par défaut, agit sur la boîte de l'utilisateur connecté (/me). La plupart des outils
acceptent un paramètre `mailbox` (adresse e-mail) pour cibler une BOÎTE PARTAGÉE/DÉLÉGUÉE
(/users/{mailbox}) à laquelle le compte a accès — nécessite les scopes `Mail.ReadWrite.Shared`
/ `Mail.Send.Shared`. Les paramètres de boîte (mailboxSettings) restent sur /me (Microsoft ne
permet pas de les modifier sur une boîte partagée en accès délégué).
"""
from __future__ import annotations

import base64
import os
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from .auth import Authenticator
from .config import Config, load_config
from .graph import GraphClient

INSTRUCTIONS = (
    "Pilote la messagerie Outlook via Microsoft Graph. Par défaut, agit sur la boîte de "
    "l'utilisateur connecté. La plupart des outils acceptent un paramètre `mailbox` "
    "(adresse e-mail) pour cibler une BOÎTE PARTAGÉE/DÉLÉGUÉE à laquelle le compte a accès "
    "(messages, dossiers, catégories, règles, envoi). Les PARAMÈTRES de boîte "
    "(get_mailbox_settings, set_automatic_replies, update_mailbox_settings) ne s'appliquent "
    "qu'à la boîte connectée — ils ne ciblent pas les boîtes partagées (limite Microsoft). "
    "Dossiers réservés : inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox. "
    "Si un outil renvoie « lancez d'abord la connexion », exécuter `login` une fois en terminal."
)

# Pour le transport HTTP (hébergé) : écoute sur toutes les interfaces, mode stateless
# (compatible scale-out / clients MCP distants). Sans effet en stdio (local).
_HTTP_PORT = int(os.environ.get("PORT") or os.environ.get("OUTLOOK_MCP_PORT") or "8000")

mcp = FastMCP(
    "outlook-graph",
    instructions=INSTRUCTIONS,
    host="0.0.0.0",
    port=_HTTP_PORT,
    stateless_http=True,
)

# --- singletons paresseux (config requise seulement à l'exécution d'un outil) ---
_cfg: Config | None = None
_auth: Authenticator | None = None
_graph: GraphClient | None = None


def init() -> GraphClient:
    global _cfg, _auth, _graph
    if _graph is None:
        _cfg = load_config()
        _auth = Authenticator(_cfg)
        _graph = GraphClient(_cfg, _auth)
    return _graph


def g() -> GraphClient:
    return _graph or init()


def get_auth() -> Authenticator:
    init()
    assert _auth is not None
    return _auth


def _base(mailbox: str | None = None) -> str:
    """Préfixe de chemin Graph : /users/{boîte} pour une boîte partagée/déléguée, sinon /me."""
    return f"/users/{mailbox}" if mailbox else "/me"


# ============================================================================
# Diagnostic
# ============================================================================
@mcp.tool()
def whoami() -> dict:
    """Vérifie la connexion : renvoie l'utilisateur courant (GET /me)."""
    return g().get("/me", params={"$select": "displayName,userPrincipalName,mail,id"})


# ============================================================================
# 📁 Dossiers
# ============================================================================
_FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount"


@mcp.tool()
def list_mail_folders(top: int = 100, include_hidden: bool = False, mailbox: str | None = None) -> list[dict]:
    """Liste les dossiers de courrier racine (GET /me|/users/{mailbox}/mailFolders).
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée)."""
    params: dict = {"$top": min(top, 100), "$select": _FOLDER_SELECT}
    if include_hidden:
        params["includeHiddenFolders"] = "true"
    return g().get_paged(f"{_base(mailbox)}/mailFolders", params=params, max_items=top)


@mcp.tool()
def list_child_folders(folder_id: str, top: int = 100, mailbox: str | None = None) -> list[dict]:
    """Liste les sous-dossiers d'un dossier (…/mailFolders/{id}/childFolders).
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée)."""
    params = {"$top": min(top, 100), "$select": _FOLDER_SELECT}
    return g().get_paged(f"{_base(mailbox)}/mailFolders/{folder_id}/childFolders", params=params, max_items=top)


@mcp.tool()
def create_mail_folder(display_name: str, parent_folder_id: str | None = None, mailbox: str | None = None) -> dict:
    """Crée un dossier (racine ou sous-dossier). parent_folder_id accepte un id ou un nom réservé.
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée)."""
    body = {"displayName": display_name}
    if parent_folder_id:
        return g().post(f"{_base(mailbox)}/mailFolders/{parent_folder_id}/childFolders", json=body)
    return g().post(f"{_base(mailbox)}/mailFolders", json=body)


@mcp.tool()
def rename_mail_folder(folder_id: str, new_name: str, mailbox: str | None = None) -> dict:
    """Renomme un dossier (PATCH …/mailFolders/{id}). mailbox : boîte partagée optionnelle."""
    return g().patch(f"{_base(mailbox)}/mailFolders/{folder_id}", json={"displayName": new_name})


@mcp.tool()
def delete_mail_folder(folder_id: str, mailbox: str | None = None) -> str:
    """Supprime un dossier ET son contenu (DELETE …/mailFolders/{id}).
    ⚠️ Action destructive et définitive — demander confirmation avant d'appeler.
    mailbox : boîte partagée optionnelle."""
    g().delete(f"{_base(mailbox)}/mailFolders/{folder_id}")
    return "Dossier supprimé."


@mcp.tool()
def list_folder_messages(
    folder_id: str,
    top: int = 25,
    search: str | None = None,
    filter: str | None = None,
    order_by: str = "receivedDateTime desc",
    mailbox: str | None = None,
) -> list[dict]:
    """Voir le contenu d'un dossier (…/mailFolders/{id}/messages).
    folder_id accepte un id ou un nom réservé. search et filter/order_by ne se combinent pas.
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée)."""
    params: dict = {
        "$top": min(top, 50),
        "$select": "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,categories,flag",
    }
    if search:
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = order_by
        if filter:
            params["$filter"] = filter
    return g().get_paged(f"{_base(mailbox)}/mailFolders/{folder_id}/messages", params=params, max_items=top)


# ============================================================================
# 🏷️ Catégories
# ============================================================================
# Couleurs Outlook : preset0..preset24. Quelques alias FR/EN courants.
COLOR_PRESETS = {
    "red": "preset0", "rouge": "preset0",
    "orange": "preset1",
    "peach": "preset2", "peche": "preset2",
    "yellow": "preset3", "jaune": "preset3",
    "green": "preset4", "vert": "preset4",
    "teal": "preset5", "turquoise": "preset5",
    "olive": "preset6",
    "blue": "preset7", "bleu": "preset7",
    "purple": "preset8", "violet": "preset8",
    "cranberry": "preset9",
    "steel": "preset10",
    "darksteel": "preset11",
    "gray": "preset12", "grey": "preset12", "gris": "preset12",
    "darkgray": "preset13",
    "black": "preset14", "noir": "preset14",
    "darkred": "preset15",
    "darkgreen": "preset19",
    "darkblue": "preset22",
}


def _color(value: str | None) -> str:
    if not value:
        return "preset0"
    v = value.strip().lower()
    if v.startswith("preset"):
        return v
    return COLOR_PRESETS.get(v, "preset0")


@mcp.tool()
def list_categories(mailbox: str | None = None) -> list[dict]:
    """Liste les catégories Outlook (…/outlook/masterCategories). mailbox : boîte partagée optionnelle."""
    return g().get_paged(f"{_base(mailbox)}/outlook/masterCategories", max_items=200)


@mcp.tool()
def create_category(display_name: str, color: str = "red", mailbox: str | None = None) -> dict:
    """Crée une catégorie avec couleur (POST …/outlook/masterCategories).
    color : nom (red, vert, bleu…) ou presetN. mailbox : boîte partagée optionnelle."""
    return g().post(
        f"{_base(mailbox)}/outlook/masterCategories",
        json={"displayName": display_name, "color": _color(color)},
    )


@mcp.tool()
def update_category(category_id: str, color: str, mailbox: str | None = None) -> dict:
    """Modifie la couleur d'une catégorie (PATCH …/outlook/masterCategories/{id}). mailbox optionnel."""
    return g().patch(f"{_base(mailbox)}/outlook/masterCategories/{category_id}", json={"color": _color(color)})


# ============================================================================
# ⚙️ Règles de boîte de réception
# ============================================================================
@mcp.tool()
def list_inbox_rules(mailbox: str | None = None) -> list[dict]:
    """Liste les règles (…/mailFolders/inbox/messageRules). mailbox : boîte partagée optionnelle."""
    return g().get_paged(f"{_base(mailbox)}/mailFolders/inbox/messageRules", max_items=200)


@mcp.tool()
def create_inbox_rule(
    display_name: str,
    actions: dict,
    conditions: dict | None = None,
    sequence: int = 1,
    is_enabled: bool = True,
    exceptions: dict | None = None,
    mailbox: str | None = None,
) -> dict:
    """Crée une règle de tri/transfert/catégorisation (POST …/mailFolders/inbox/messageRules).

    conditions / actions / exceptions = objets Graph (messageRulePredicates / messageRuleActions).
    Exemple actions : {"moveToFolder": "<folderId>", "markAsRead": true,
                       "assignCategories": ["Travail"], "stopProcessingRules": true}
    Exemple conditions : {"senderContains": ["@client.com"]}  ou  {"subjectContains": ["Facture"]}.
    mailbox : boîte partagée optionnelle (peut nécessiter des droits étendus).
    """
    body: dict = {
        "displayName": display_name,
        "sequence": sequence,
        "isEnabled": is_enabled,
        "actions": actions,
    }
    if conditions:
        body["conditions"] = conditions
    if exceptions:
        body["exceptions"] = exceptions
    return g().post(f"{_base(mailbox)}/mailFolders/inbox/messageRules", json=body)


@mcp.tool()
def update_inbox_rule(rule_id: str, changes: dict, mailbox: str | None = None) -> dict:
    """Modifie / active / désactive une règle (PATCH …/messageRules/{id}).
    changes : displayName, isEnabled, sequence, conditions, actions, exceptions.
    mailbox : boîte partagée optionnelle."""
    return g().patch(f"{_base(mailbox)}/mailFolders/inbox/messageRules/{rule_id}", json=changes)


@mcp.tool()
def delete_inbox_rule(rule_id: str, mailbox: str | None = None) -> str:
    """Supprime une règle (DELETE …/messageRules/{id}). mailbox : boîte partagée optionnelle."""
    g().delete(f"{_base(mailbox)}/mailFolders/inbox/messageRules/{rule_id}")
    return "Règle supprimée."


# ============================================================================
# ✉️ Messages
# ============================================================================
def _recipients(addresses: list[str] | None) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in (addresses or [])]


def _file_attachment(path_str: str, name: str | None = None, content_type: str | None = None) -> dict:
    p = Path(path_str)
    data = p.read_bytes()
    if len(data) > 3 * 1024 * 1024:
        raise ValueError(
            f"« {p.name} » > 3 Mo : une pièce jointe directe est limitée à 3 Mo "
            "(les gros fichiers nécessitent une session d'upload, non gérée ici)."
        )
    att = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name or p.name,
        "contentBytes": base64.b64encode(data).decode(),
    }
    if content_type:
        att["contentType"] = content_type
    return att


@mcp.tool()
def list_messages(
    top: int = 25,
    search: str | None = None,
    filter: str | None = None,
    order_by: str = "receivedDateTime desc",
    select: str | None = None,
    mailbox: str | None = None,
) -> list[dict]:
    """Rechercher / lister des messages (…/messages).
    search = recherche plein texte. filter = expression $filter OData (ne se combine pas avec search).
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée)."""
    params: dict = {"$top": min(top, 50)}
    params["$select"] = select or (
        "id,subject,from,toRecipients,receivedDateTime,isRead,importance,"
        "hasAttachments,categories,flag,bodyPreview"
    )
    if search:
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = order_by
        if filter:
            params["$filter"] = filter
    return g().get_paged(f"{_base(mailbox)}/messages", params=params, max_items=top)


@mcp.tool()
def get_message(
    message_id: str, include_body: bool = True, include_attachments: bool = False, mailbox: str | None = None
) -> dict:
    """Lire un message complet (…/messages/{id}).
    include_attachments=True développe les pièces jointes (base64). mailbox : boîte partagée optionnelle."""
    params: dict = {}
    if include_attachments:
        params["$expand"] = "attachments"
    if not include_body:
        params["$select"] = (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,"
            "importance,categories,flag,hasAttachments,bodyPreview"
        )
    return g().get(f"{_base(mailbox)}/messages/{message_id}", params=params or None)


@mcp.tool()
def update_message(
    message_id: str,
    is_read: bool | None = None,
    importance: str | None = None,
    categories: list[str] | None = None,
    flag_status: str | None = None,
    mailbox: str | None = None,
) -> dict:
    """Marquer lu/non lu, importance, classement (catégories), drapeau (PATCH …/messages/{id}).
    importance : low|normal|high. flag_status : notFlagged|flagged|complete.
    mailbox : boîte partagée optionnelle."""
    body: dict = {}
    if is_read is not None:
        body["isRead"] = is_read
    if importance is not None:
        body["importance"] = importance
    if categories is not None:
        body["categories"] = categories
    if flag_status is not None:
        body["flag"] = {"flagStatus": flag_status}
    if not body:
        raise ValueError("Aucune modification fournie.")
    return g().patch(f"{_base(mailbox)}/messages/{message_id}", json=body)


@mcp.tool()
def move_message(message_id: str, destination_folder_id: str, mailbox: str | None = None) -> dict:
    """Déplace un message vers un dossier (POST …/messages/{id}/move).
    destination_folder_id : id ou nom réservé. mailbox : boîte partagée optionnelle."""
    return g().post(f"{_base(mailbox)}/messages/{message_id}/move", json={"destinationId": destination_folder_id})


@mcp.tool()
def delete_message(message_id: str, mailbox: str | None = None) -> str:
    """Supprime un message → Éléments supprimés (DELETE …/messages/{id}). mailbox : boîte partagée optionnelle."""
    g().delete(f"{_base(mailbox)}/messages/{message_id}")
    return "Message déplacé vers les Éléments supprimés."


@mcp.tool()
def send_mail(
    to: list[str],
    subject: str,
    body: str,
    body_type: str = "HTML",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[str] | None = None,
    save_to_sent: bool = True,
    mailbox: str | None = None,
) -> str:
    """Envoie un e-mail (POST …/sendMail). body_type : HTML ou Text. attachments = chemins locaux (≤3 Mo).
    mailbox : envoie AU NOM d'une boîte partagée (nécessite Mail.Send.Shared + droits SendAs)."""
    message: dict = {
        "subject": subject,
        "body": {"contentType": body_type, "content": body},
        "toRecipients": _recipients(to),
    }
    if cc:
        message["ccRecipients"] = _recipients(cc)
    if bcc:
        message["bccRecipients"] = _recipients(bcc)
    if attachments:
        message["attachments"] = [_file_attachment(p) for p in attachments]
    g().post(f"{_base(mailbox)}/sendMail", json={"message": message, "saveToSentItems": save_to_sent}, expect_json=False)
    return "E-mail envoyé."


@mcp.tool()
def reply_message(message_id: str, comment: str, reply_all: bool = False, mailbox: str | None = None) -> str:
    """Répondre / Répondre à tous (POST …/messages/{id}/reply | /replyAll). mailbox : boîte partagée optionnelle."""
    endpoint = "replyAll" if reply_all else "reply"
    g().post(f"{_base(mailbox)}/messages/{message_id}/{endpoint}", json={"comment": comment}, expect_json=False)
    return "Réponse envoyée."


@mcp.tool()
def forward_message(message_id: str, to: list[str], comment: str = "", mailbox: str | None = None) -> str:
    """Transférer un message (POST …/messages/{id}/forward). mailbox : boîte partagée optionnelle."""
    g().post(
        f"{_base(mailbox)}/messages/{message_id}/forward",
        json={"comment": comment, "toRecipients": _recipients(to)},
        expect_json=False,
    )
    return "Message transféré."


@mcp.tool()
def create_reply_draft(
    message_id: str,
    comment: str = "",
    reply_all: bool = False,
    mailbox: str | None = None,
) -> dict:
    """Crée un BROUILLON de réponse rattaché au fil d'origine (POST …/messages/{id}/createReply
    ou /createReplyAll) — même conversation, objet « Re: … », historique cité et destinataires
    repris — SANS l'envoyer. Le brouillon va dans les Brouillons : on le relit et on l'envoie
    soi-même depuis Outlook (aucun risque d'envoi involontaire). Renvoie l'objet brouillon.

    comment : texte de la réponse, inséré AU-DESSUS de l'historique cité (optionnel).
    reply_all : True = répondre à tous (createReplyAll) ; False = à l'expéditeur seul (createReply).
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée).
    Pour réécrire entièrement le corps HTML, créez le brouillon puis éditez-le avec update_draft
    (attention : update_draft remplace le corps, donc l'historique cité). Pour ajouter du texte en
    gardant l'historique, passez plutôt par `comment` ici.
    """
    endpoint = "createReplyAll" if reply_all else "createReply"
    body = {"comment": comment} if comment else None
    return g().post(f"{_base(mailbox)}/messages/{message_id}/{endpoint}", json=body)


@mcp.tool()
def create_forward_draft(
    message_id: str,
    to: list[str] | None = None,
    comment: str = "",
    mailbox: str | None = None,
) -> dict:
    """Crée un BROUILLON de transfert du message (POST …/messages/{id}/createForward) — contenu
    d'origine repris — SANS l'envoyer. Le brouillon va dans les Brouillons : on le relit et on
    l'envoie soi-même depuis Outlook. Renvoie l'objet brouillon.

    to : destinataires du transfert (optionnel — complétables ensuite dans Outlook).
    comment : texte ajouté au-dessus du message transféré (optionnel).
    mailbox : adresse d'une boîte partagée (sinon la boîte connectée).
    """
    body: dict = {}
    if comment:
        body["comment"] = comment
    if to:
        body["toRecipients"] = _recipients(to)
    return g().post(f"{_base(mailbox)}/messages/{message_id}/createForward", json=body or None)


@mcp.tool()
def create_draft(
    to: list[str] | None = None,
    subject: str = "",
    body: str = "",
    body_type: str = "HTML",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    mailbox: str | None = None,
) -> dict:
    """Crée un brouillon (POST …/messages). Renvoie l'objet message. mailbox : boîte partagée optionnelle."""
    msg: dict = {"subject": subject, "body": {"contentType": body_type, "content": body}}
    if to:
        msg["toRecipients"] = _recipients(to)
    if cc:
        msg["ccRecipients"] = _recipients(cc)
    if bcc:
        msg["bccRecipients"] = _recipients(bcc)
    return g().post(f"{_base(mailbox)}/messages", json=msg)


@mcp.tool()
def update_draft(message_id: str, changes: dict, mailbox: str | None = None) -> dict:
    """Modifie un brouillon (PATCH …/messages/{id}). changes au format Graph. mailbox : boîte partagée optionnelle."""
    return g().patch(f"{_base(mailbox)}/messages/{message_id}", json=changes)


@mcp.tool()
def send_draft(message_id: str, mailbox: str | None = None) -> str:
    """Envoie un brouillon existant (POST …/messages/{id}/send). mailbox : boîte partagée optionnelle."""
    g().post(f"{_base(mailbox)}/messages/{message_id}/send", expect_json=False)
    return "Brouillon envoyé."


@mcp.tool()
def add_attachment(message_id: str, file_path: str, name: str | None = None, mailbox: str | None = None) -> dict:
    """Ajoute une pièce jointe à un message/brouillon (POST …/messages/{id}/attachments).
    Fichier ≤ 3 Mo. mailbox : boîte partagée optionnelle."""
    return g().post(f"{_base(mailbox)}/messages/{message_id}/attachments", json=_file_attachment(file_path, name))


# ============================================================================
# 🔧 Paramètres de la boîte — UNIQUEMENT la boîte connectée (/me)
# ============================================================================
# Microsoft ne permet pas de modifier les mailboxSettings d'une boîte partagée en accès
# délégué (pas de scope MailboxSettings.Shared) → ces outils restent sur /me.
@mcp.tool()
def get_mailbox_settings() -> dict:
    """Lit les paramètres de VOTRE boîte (GET /me/mailboxSettings) :
    fuseau, langue, formats, réponses automatiques, heures de travail. (Boîte connectée uniquement.)"""
    return g().get("/me/mailboxSettings")


@mcp.tool()
def set_automatic_replies(
    status: str,
    internal_message: str = "",
    external_message: str | None = None,
    external_audience: str = "all",
    start: str | None = None,
    end: str | None = None,
    time_zone: str = "Romance Standard Time",
) -> dict:
    """Réponses automatiques / Absence du bureau (PATCH /me/mailboxSettings). BOÎTE CONNECTÉE uniquement.
    status : disabled | alwaysEnabled | scheduled. Pour 'scheduled', fournir start et end (ISO).
    external_audience : none | contactsOnly | all. time_zone : nom Windows (ex. 'Romance Standard Time')."""
    setting: dict = {
        "status": status,
        "externalAudience": external_audience,
        "internalReplyMessage": internal_message,
        "externalReplyMessage": external_message if external_message is not None else internal_message,
    }
    if status == "scheduled":
        if not (start and end):
            raise ValueError("start et end sont requis pour status='scheduled'.")
        setting["scheduledStartDateTime"] = {"dateTime": start, "timeZone": time_zone}
        setting["scheduledEndDateTime"] = {"dateTime": end, "timeZone": time_zone}
    return g().patch("/me/mailboxSettings", json={"automaticRepliesSetting": setting})


@mcp.tool()
def update_mailbox_settings(
    time_zone: str | None = None,
    language_locale: str | None = None,
    date_format: str | None = None,
    time_format: str | None = None,
) -> dict:
    """Fuseau horaire, langue, format de date/heure (PATCH /me/mailboxSettings). BOÎTE CONNECTÉE uniquement.
    Ex. time_zone='Romance Standard Time', language_locale='fr-FR', date_format='dd/MM/yyyy', time_format='HH:mm'."""
    body: dict = {}
    if time_zone is not None:
        body["timeZone"] = time_zone
    if language_locale is not None:
        body["language"] = {"locale": language_locale}
    if date_format is not None:
        body["dateFormat"] = date_format
    if time_format is not None:
        body["timeFormat"] = time_format
    if not body:
        raise ValueError("Aucun paramètre fourni.")
    return g().patch("/me/mailboxSettings", json=body)
