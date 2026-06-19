"""Client HTTP léger pour Microsoft Graph."""
from __future__ import annotations

from typing import Any

import httpx

from .auth import Authenticator
from .config import GRAPH_BASE, Config


class GraphError(RuntimeError):
    """Erreur renvoyée par Microsoft Graph, message lisible."""


class GraphClient:
    def __init__(self, cfg: Config, auth: Authenticator):
        self.cfg = cfg
        self.auth = auth
        self._client = httpx.Client(base_url=GRAPH_BASE, timeout=60.0)

    def _headers(self, extra: dict | None = None) -> dict:
        headers = {
            "Authorization": f"Bearer {self.auth.get_token()}",
            "Content-Type": "application/json",
        }
        if self.cfg.immutable_ids:
            headers["Prefer"] = 'IdType="ImmutableId"'
        if extra:
            headers.update(extra)
        return headers

    def request(
        self,
        method: str,
        path: str,
        *,
        params: dict | None = None,
        json: Any = None,
        headers: dict | None = None,
        expect_json: bool = True,
    ) -> Any:
        # path relatif ('/me/...') ou absolu (nextLink) : httpx gère les deux.
        resp = self._client.request(
            method, path, params=params, json=json, headers=self._headers(headers)
        )
        if resp.status_code >= 400:
            self._raise(resp)
        if not expect_json or resp.status_code == 204 or not resp.content:
            return None
        return resp.json()

    @staticmethod
    def _raise(resp: httpx.Response) -> None:
        try:
            err = resp.json().get("error", {})
            detail = f"{err.get('code', resp.status_code)} — {err.get('message', resp.text)}"
        except Exception:
            detail = f"{resp.status_code} — {resp.text}"
        raise GraphError(f"Erreur Graph {resp.status_code} ({resp.request.method} {resp.request.url})\n{detail}")

    def get(self, path: str, **kw) -> Any:
        return self.request("GET", path, **kw)

    def post(self, path: str, **kw) -> Any:
        return self.request("POST", path, **kw)

    def patch(self, path: str, **kw) -> Any:
        return self.request("PATCH", path, **kw)

    def delete(self, path: str, **kw) -> Any:
        kw.setdefault("expect_json", False)
        return self.request("DELETE", path, **kw)

    def get_paged(self, path: str, *, params: dict | None = None, max_items: int = 50) -> list[dict]:
        """Suit @odata.nextLink jusqu'à max_items."""
        items: list[dict] = []
        data = self.get(path, params=params)
        while data:
            items.extend(data.get("value", []))
            if len(items) >= max_items:
                return items[:max_items]
            nxt = data.get("@odata.nextLink")
            if not nxt:
                break
            data = self.get(nxt)  # URL absolue, déjà paramétrée
        return items
