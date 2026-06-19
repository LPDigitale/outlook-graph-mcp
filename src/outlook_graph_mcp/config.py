"""Configuration lue depuis l'environnement."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Permissions déléguées requises (consenties au login).
DEFAULT_SCOPES = ["Mail.ReadWrite", "Mail.Send", "MailboxSettings.ReadWrite"]


def _truthy(value: str | None) -> bool:
    return (value or "").strip().lower() in ("1", "true", "yes", "on", "oui")


def _cache_dir() -> Path:
    override = os.environ.get("OUTLOOK_MCP_CACHE_DIR")
    if override:
        return Path(override).expanduser()
    return Path.home() / ".outlook-graph-mcp"


@dataclass
class Config:
    client_id: str
    tenant_id: str
    scopes: list[str]
    cache_dir: Path
    immutable_ids: bool
    # Backend de stockage du cache de jeton : 'local' (fichier chiffré) ou 'keyvault' (Azure).
    cache_backend: str
    keyvault_url: str | None
    cache_secret_name: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    @property
    def cache_file(self) -> Path:
        return self.cache_dir / "token_cache.bin"


def load_config() -> Config:
    client_id = os.environ.get("OUTLOOK_MCP_CLIENT_ID", "").strip()
    if not client_id:
        raise RuntimeError(
            "OUTLOOK_MCP_CLIENT_ID manquant. Renseignez le « Application (client) ID » "
            "de votre inscription d'application Azure AD (voir README, étape A)."
        )

    # Compte professionnel mono-tenant : mettre le Tenant ID (GUID) ou le domaine.
    tenant_id = os.environ.get("OUTLOOK_MCP_TENANT_ID", "").strip() or "organizations"

    scopes_env = os.environ.get("OUTLOOK_MCP_SCOPES")
    scopes = scopes_env.split() if scopes_env else list(DEFAULT_SCOPES)

    immutable_ids = _truthy(os.environ.get("OUTLOOK_MCP_IMMUTABLE_IDS", "true"))

    cache_backend = (os.environ.get("OUTLOOK_MCP_CACHE_BACKEND", "local").strip().lower() or "local")
    keyvault_url = os.environ.get("OUTLOOK_MCP_KEYVAULT_URL", "").strip() or None
    cache_secret_name = os.environ.get("OUTLOOK_MCP_CACHE_SECRET", "outlook-token-cache").strip()

    if cache_backend == "keyvault" and not keyvault_url:
        raise RuntimeError(
            "OUTLOOK_MCP_CACHE_BACKEND=keyvault mais OUTLOOK_MCP_KEYVAULT_URL est vide."
        )

    cache_dir = _cache_dir()
    cache_dir.mkdir(parents=True, exist_ok=True)

    return Config(
        client_id=client_id,
        tenant_id=tenant_id,
        scopes=scopes,
        cache_dir=cache_dir,
        immutable_ids=immutable_ids,
        cache_backend=cache_backend,
        keyvault_url=keyvault_url,
        cache_secret_name=cache_secret_name,
    )
