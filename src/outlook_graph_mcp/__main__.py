"""Point d'entrée.

Usage :
    outlook-graph-mcp                 # serveur MCP local (stdio) — lancé par Claude Desktop/Code
    outlook-graph-mcp serve-http      # serveur MCP HTTP (hébergé / conteneur)
    outlook-graph-mcp login           # connexion Microsoft (device-code) — à faire une fois
    outlook-graph-mcp login --interactive   # connexion via navigateur
    outlook-graph-mcp whoami          # teste la connexion
    outlook-graph-mcp logout          # vide le cache de jeton

Le backend de cache (local chiffré vs Azure Key Vault) est choisi par
OUTLOOK_MCP_CACHE_BACKEND. Pour amorcer le serveur hébergé, lancez `login` en local
avec OUTLOOK_MCP_CACHE_BACKEND=keyvault et OUTLOOK_MCP_KEYVAULT_URL renseignés.
"""
from __future__ import annotations

import os
import sys

USAGE = __doc__


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    cmd = args[0] if args else "serve"

    if cmd in ("-h", "--help", "help"):
        print(USAGE)
        return 0

    if cmd == "login":
        auth = _get_auth()
        if "--interactive" in args:
            auth.login_interactive()
        else:
            auth.login_device_code()
        _print_me("Connecté")
        return 0

    if cmd == "logout":
        _get_auth().logout()
        print("Déconnecté (cache de jeton vidé).")
        return 0

    if cmd in ("whoami", "test"):
        _print_me("Connecté")
        return 0

    if cmd == "serve-http":
        import uvicorn

        from .webapp import build_http_app

        app = build_http_app()
        port = int(os.environ.get("PORT") or os.environ.get("OUTLOOK_MCP_PORT") or "8000")
        uvicorn.run(app, host="0.0.0.0", port=port)
        return 0

    if cmd == "serve":
        from .server import init, mcp

        init()  # valide la config et prépare l'auth dès le démarrage
        mcp.run()  # stdio
        return 0

    print(f"Commande inconnue : {cmd}\n", file=sys.stderr)
    print(USAGE, file=sys.stderr)
    return 2


def _get_auth():
    from .server import get_auth

    return get_auth()


def _print_me(prefix: str) -> None:
    from .server import g

    me = g().get("/me", params={"$select": "displayName,userPrincipalName,mail"})
    print(f"{prefix} : {me.get('displayName')} <{me.get('userPrincipalName')}>")


if __name__ == "__main__":
    raise SystemExit(main())
