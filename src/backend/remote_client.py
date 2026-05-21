from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv


def _runtime_root() -> Path:
    """Retorna a raiz usada em desenvolvimento ou no executável."""
    import sys

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    # remote_client.py -> .../LegalHub/src/backend/remote_client.py
    return Path(__file__).resolve().parents[2]


ROOT = _runtime_root()
load_dotenv(ROOT / ".env", override=False)


class RemoteApiError(RuntimeError):
    """Erro de comunicação ou resposta inválida da API remota."""


class RemoteClient:
    """Cliente HTTP para a API hospedada na VM.

    Importante:
    - O app desktop nunca deve receber chaves de IA.
    - As chaves reais ficam apenas na VM.
    - Este cliente envia arquivos/dados e recebe o resultado processado.
    """

    def __init__(self, base_url: str | None = None, token: str | None = None, timeout: int = 180) -> None:
        self.base_url = (base_url or os.getenv("LEGALHUB_API_BASE_URL") or "").strip().rstrip("/")
        self.token = (token or os.getenv("LEGALHUB_CLIENT_TOKEN") or "").strip()
        self.timeout = int(os.getenv("LEGALHUB_API_TIMEOUT", str(timeout)))

    @property
    def enabled(self) -> bool:
        return bool(self.base_url)

    def _headers(self) -> dict[str, str]:
        headers = {"Accept": "application/json"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    def get(self, path: str) -> Any:
        return self._request("GET", path)

    def post_json(self, path: str, payload: dict[str, Any]) -> Any:
        return self._request("POST", path, json=payload)

    def post_files(
        self,
        path: str,
        files: list[Path],
        data: dict[str, Any] | None = None,
        file_field: str = "files",
    ) -> Any:
        if not self.enabled:
            raise RemoteApiError("LEGALHUB_API_BASE_URL não configurada no .env local.")

        opened_files: list[Any] = []
        try:
            multipart = []
            for file_path in files:
                p = Path(file_path)
                if not p.exists() or not p.is_file():
                    raise FileNotFoundError(f"Arquivo não encontrado: {p}")
                handle = open(p, "rb")
                opened_files.append(handle)
                multipart.append((file_field, (p.name, handle, "application/pdf")))

            response = requests.post(
                self._url(path),
                headers=self._headers(),
                files=multipart,
                data=data or {},
                timeout=self.timeout,
            )
            return self._parse_response(response)
        finally:
            for handle in opened_files:
                try:
                    handle.close()
                except Exception:
                    pass

    def _request(self, method: str, path: str, **kwargs: Any) -> Any:
        if not self.enabled:
            raise RemoteApiError("LEGALHUB_API_BASE_URL não configurada no .env local.")

        response = requests.request(
            method,
            self._url(path),
            headers=self._headers(),
            timeout=self.timeout,
            **kwargs,
        )
        return self._parse_response(response)

    def _url(self, path: str) -> str:
        normalized = path if path.startswith("/") else f"/{path}"
        return f"{self.base_url}{normalized}"

    @staticmethod
    def _parse_response(response: requests.Response) -> Any:
        try:
            payload = response.json()
        except ValueError:
            payload = response.text

        if response.status_code >= 400:
            raise RemoteApiError(f"Erro HTTP {response.status_code}: {payload}")

        return payload


remote_client = RemoteClient()
