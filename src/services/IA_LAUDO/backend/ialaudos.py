# src/services/IA_LAUDO/backend/main.py
from __future__ import annotations
import os, sys
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI

_client: OpenAI | None = None

def _app_base() -> Path:
    # PyInstaller
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys.executable).resolve().parent
    # raiz do projeto (…/src/services/IA_LAUDO/backend/main.py -> parents[4])
    return Path(__file__).resolve().parents[4]

env_path = _app_base() / ".env"
load_dotenv(dotenv_path=env_path, override=False)

_PROMPT_ID = (os.getenv("OPENAI_PROMPT_ID") or "").strip()

def _get_client() -> OpenAI:
    global _client
    if _client is None:
        api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
        if not api_key:
            raise RuntimeError(
                "OPENAI_API_KEY não encontrada.\n"
                f"Procurei em: {env_path}\n"
                "Crie/ajuste o .env na raiz do projeto com:\n"
                "OPENAI_API_KEY=sua_chave_aqui"
            )
        _client = OpenAI(api_key=api_key)
    return _client

def analisar_pdf(caminho_pdf: Path) -> str:
    try:
        if not caminho_pdf.exists() or not caminho_pdf.is_file():
            return "Erro: arquivo não encontrado."
        if caminho_pdf.suffix.lower() != ".pdf":
            return "Erro: este assistente aceita apenas arquivos PDF."
        if not _PROMPT_ID:
            return "Erro: OPENAI_PROMPT_ID não configurado no .env."

        client = _get_client()

        with open(caminho_pdf, "rb") as f:
            uploaded_file = client.files.create(file=f, purpose="assistants")

        file_id = uploaded_file.id

        response = client.responses.create(
            prompt={"id": _PROMPT_ID},
            input=[{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": "Analise o laudo em anexo."},
                    {"type": "input_file", "file_id": file_id},
                ],
            }],
        )

        saida = (response.output_text or "").strip()
        return saida or "(Sem conteúdo retornado pelo agente.)"

    except Exception as exc:
        return f"Erro ao consultar a IA: {exc!s}"

    finally:
        try:
            if "file_id" in locals() and file_id and _client:
                _client.files.delete(file_id)
        except Exception:
            pass
