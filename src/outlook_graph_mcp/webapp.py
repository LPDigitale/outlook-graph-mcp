"""Application HTTP (transport streamable-http) pour l'hébergement distant.

Expose le serveur MCP sur /mcp, protégé par une clé d'API (en-tête `x-api-key`,
paramètre `?key=`, ou `Authorization: Bearer …`). /healthz reste public (sondes).

L'authentification est un middleware ASGI pur (pas BaseHTTPMiddleware) afin de ne
pas bufferiser le flux SSE du protocole MCP.
"""
from __future__ import annotations

import os
from urllib.parse import parse_qs

_UNAUTHORIZED = b'{"error":"unauthorized"}'


class ApiKeyMiddleware:
    def __init__(self, app, api_key: str, public_paths: set[str]):
        self.app = app
        self.api_key = api_key
        self.public_paths = public_paths

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)

        path = scope.get("path", "")
        method = scope.get("method", "GET")
        if path in self.public_paths or method == "OPTIONS":
            return await self.app(scope, receive, send)

        headers = {k.decode().lower(): v.decode() for k, v in scope.get("headers", [])}
        provided = headers.get("x-api-key")
        if not provided:
            provided = parse_qs(scope.get("query_string", b"").decode()).get("key", [None])[0]
        if not provided:
            auth = headers.get("authorization", "")
            if auth.lower().startswith("bearer "):
                provided = auth[7:]

        if provided != self.api_key:
            await send(
                {
                    "type": "http.response.start",
                    "status": 401,
                    "headers": [(b"content-type", b"application/json")],
                }
            )
            await send({"type": "http.response.body", "body": _UNAUTHORIZED})
            return

        return await self.app(scope, receive, send)


async def _healthz(request):
    from starlette.responses import PlainTextResponse

    return PlainTextResponse("ok")


def build_http_app():
    """Construit l'app ASGI : valide la config/auth, monte /healthz, protège par clé."""
    from .server import init, mcp

    init()  # valide la configuration et prépare l'auth dès le démarrage
    app = mcp.streamable_http_app()
    app.add_route("/healthz", _healthz, methods=["GET"])

    api_key = os.environ.get("OUTLOOK_MCP_API_KEY")
    if api_key:
        app.add_middleware(ApiKeyMiddleware, api_key=api_key, public_paths={"/healthz"})
    return app
