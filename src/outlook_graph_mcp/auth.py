"""Authentification MSAL (public client) avec cache de jeton persistant.

Deux backends de cache, choisis par `OUTLOOK_MCP_CACHE_BACKEND` :
- ``local``    : fichier chiffré (DPAPI sous Windows via msal-extensions). Usage poste de travail.
- ``keyvault`` : un secret Azure Key Vault. Usage serveur hébergé (identité managée).

La connexion device-code se fait une fois (`login`) ; le serveur ne fait ensuite que des
acquisitions silencieuses (refresh token). Le login peut être exécuté en local en pointant
sur le Key Vault pour amorcer le serveur distant.
"""
from __future__ import annotations

import threading
from typing import Callable

import msal

from .config import Config

try:  # cache local chiffré (DPAPI Windows)
    from msal_extensions import (
        FilePersistence,
        PersistedTokenCache,
        build_encrypted_persistence,
    )

    _HAVE_EXT = True
except Exception:  # pragma: no cover
    _HAVE_EXT = False


class KeyVaultStore:
    """Lit/écrit le cache MSAL (sérialisé) dans un secret Azure Key Vault."""

    def __init__(self, vault_url: str, secret_name: str):
        from azure.identity import DefaultAzureCredential
        from azure.keyvault.secrets import SecretClient

        self._secret_name = secret_name
        self._client = SecretClient(vault_url=vault_url, credential=DefaultAzureCredential())

    def load(self) -> str | None:
        try:
            return self._client.get_secret(self._secret_name).value
        except Exception:
            # Secret absent (1er amorçage) ou accès indisponible : cache vide.
            return None

    def save(self, data: str) -> None:
        self._client.set_secret(self._secret_name, data)


def _build_cache(cfg: Config):
    """Retourne (cache, persist) où persist est un callable(serialized)->None ou None (auto)."""
    if cfg.cache_backend == "keyvault":
        store = KeyVaultStore(cfg.keyvault_url or "", cfg.cache_secret_name)
        cache = msal.SerializableTokenCache()
        blob = store.load()
        if blob:
            cache.deserialize(blob)
        return cache, store.save

    # backend local
    if _HAVE_EXT:
        location = str(cfg.cache_file)
        try:
            persistence = build_encrypted_persistence(location)
        except Exception:
            persistence = FilePersistence(location)
        return PersistedTokenCache(persistence), None  # persistance automatique

    cache = msal.SerializableTokenCache()
    if cfg.cache_file.exists():
        try:
            cache.deserialize(cfg.cache_file.read_text(encoding="utf-8"))
        except Exception:
            pass

    def _file_persist(data: str) -> None:
        cfg.cache_file.write_text(data, encoding="utf-8")

    return cache, _file_persist


class Authenticator:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self._cache, self._persist = _build_cache(cfg)
        self._app = msal.PublicClientApplication(
            cfg.client_id,
            authority=cfg.authority,
            token_cache=self._cache,
        )
        self._lock = threading.Lock()

    def _save(self) -> None:
        if self._persist is not None and getattr(self._cache, "has_state_changed", False):
            self._persist(self._cache.serialize())

    # -- acquisition silencieuse pour les appels Graph ----------------------------
    def get_token(self) -> str:
        with self._lock:
            result = None
            accounts = self._app.get_accounts()
            if accounts:
                result = self._app.acquire_token_silent(self.cfg.scopes, account=accounts[0])
            self._save()
            if not result or "access_token" not in result:
                raise RuntimeError(
                    "Aucun jeton valide en cache. Connectez-vous d'abord (commande `login`) — "
                    "en local pour un usage poste, ou en pointant sur le Key Vault pour amorcer "
                    "le serveur hébergé."
                )
            return result["access_token"]

    # -- connexions interactives (CLI) -------------------------------------------
    def login_device_code(self, print_fn: Callable[[str], None] = print) -> dict:
        with self._lock:
            flow = self._app.initiate_device_flow(scopes=self.cfg.scopes)
            if "user_code" not in flow:
                raise RuntimeError(
                    "Échec du device flow. Vérifiez « Allow public client flows » sur l'app Azure. "
                    f"Détail : {flow.get('error_description', flow)}"
                )
            print_fn(flow["message"])
            result = self._app.acquire_token_by_device_flow(flow)  # bloque jusqu'à validation
            self._save()
        if "access_token" not in result:
            raise RuntimeError(f"Connexion échouée : {result.get('error_description', result)}")
        return result

    def login_interactive(self, print_fn: Callable[[str], None] = print) -> dict:
        with self._lock:
            print_fn("Ouverture du navigateur pour la connexion Microsoft…")
            result = self._app.acquire_token_interactive(scopes=self.cfg.scopes)
            self._save()
        if "access_token" not in result:
            raise RuntimeError(f"Connexion échouée : {result.get('error_description', result)}")
        return result

    def logout(self) -> None:
        with self._lock:
            for acc in self._app.get_accounts():
                self._app.remove_account(acc)
            # Force une réécriture du cache vidé.
            if self._persist is not None:
                self._persist(self._cache.serialize())
