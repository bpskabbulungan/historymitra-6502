from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any

import requests

DEFAULT_BASE_URL = "https://mitra-api.bps.go.id"


class MitraApiError(RuntimeError):
    pass


@dataclass(slots=True)
class MitraApiConfig:
    base_url: str = DEFAULT_BASE_URL
    cookie: str | None = None
    authorization: str | None = None
    headers_json: str | None = None
    user_agent: str | None = None
    referer: str | None = None
    timeout_seconds: int = 30

    @classmethod
    def from_env(cls) -> "MitraApiConfig":
        return cls(
            base_url=os.getenv("MITRA_API_BASE_URL", DEFAULT_BASE_URL),
            cookie=os.getenv("MITRA_API_COOKIE"),
            authorization=os.getenv("MITRA_API_AUTHORIZATION"),
            headers_json=os.getenv("MITRA_API_HEADERS_JSON"),
            user_agent=os.getenv("MITRA_API_USER_AGENT"),
            referer=os.getenv("MITRA_API_REFERER"),
            timeout_seconds=int(os.getenv("MITRA_API_TIMEOUT_SECONDS", "30")),
        )

    def build_headers(self) -> dict[str, str]:
        headers: dict[str, str] = {
            "Accept": "application/json, text/plain, */*",
        }
        if self.user_agent:
            headers["User-Agent"] = self.user_agent
        if self.referer:
            headers["Referer"] = self.referer
        if self.cookie:
            headers["Cookie"] = self.cookie
        if self.authorization:
            headers["Authorization"] = self.authorization
        if self.headers_json:
            extra_headers = json.loads(self.headers_json)
            for key, value in extra_headers.items():
                if value is not None:
                    headers[str(key)] = str(value)
        return headers


class MitraApiClient:
    def __init__(self, config: MitraApiConfig) -> None:
        self.config = config
        self.session = requests.Session()
        self.session.headers.update(config.build_headers())

    def _get(self, path: str, *, params: dict[str, Any] | None = None) -> Any:
        url = f"{self.config.base_url.rstrip('/')}/{path.lstrip('/')}"
        response = self.session.get(url, params=params, timeout=self.config.timeout_seconds)
        if response.status_code == 401:
            raise MitraApiError(
                "API mengembalikan 401 Unauthorized. Isi autentikasi lewat "
                "MITRA_API_COOKIE, MITRA_API_AUTHORIZATION, atau MITRA_API_HEADERS_JSON."
            )
        if not response.ok:
            raise MitraApiError(
                f"Gagal mengambil {url} dengan status {response.status_code}: {response.text[:300]}"
            )
        try:
            return response.json()
        except requests.JSONDecodeError as exc:
            raise MitraApiError(f"Respons dari {url} bukan JSON yang valid.") from exc

    def fetch_mitra_list(self, *, year: int, prov: str, kab: str) -> list[dict[str, Any]]:
        payload = self._get(f"/api/mitra-kepka/by-year-wil/{year}/{prov}/{kab}")
        return list(payload.get("mitras", []))

    def fetch_mitra_detail(self, *, id_mitra: str) -> dict[str, Any]:
        payload = self._get(f"/api/mitra/id/{id_mitra}")
        return dict(payload.get("mitra", {}))

    def fetch_mitra_history_documents(
        self, *, id_mitra: str, year: int, prev: bool = True
    ) -> list[dict[str, Any]]:
        url = f"{self.config.base_url.rstrip('/')}/api/mitra/hist/sm/{id_mitra}"
        response = self.session.get(
            url,
            params={"tahun": str(year), "prev": str(prev).lower()},
            timeout=self.config.timeout_seconds,
        )
        if response.status_code == 401:
            raise MitraApiError(
                "API histori mengembalikan 401 Unauthorized. Pastikan header atau cookie sesi sudah valid."
            )
        if not response.ok:
            raise MitraApiError(
                f"Gagal mengambil histori {id_mitra} dengan status {response.status_code}: {response.text[:300]}"
            )
        from .parsers import parse_multi_json_documents

        return parse_multi_json_documents(response.text)
