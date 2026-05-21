from __future__ import annotations

from dataclasses import asdict
from typing import Any

from .lotes_service import (
    gerar_lotes,
    info_modelo_padrao,
    listar_colunas_lote,
)


class GeradorLotesAPI:
    """
    Métodos retornam apenas tipos serializáveis (dict/list/str/etc),
    para funcionar bem com window.pywebview.api.*
    """

    def get_default_model(self) -> dict[str, str]:
        return info_modelo_padrao()

    def preview_columns(self, caminho_lote: str, coluna_processo_hint: str | None = None) -> dict[str, Any]:
        cols = listar_colunas_lote(caminho_lote, coluna_processo_hint=coluna_processo_hint)
        return {"columns": cols}

    def generate(self, payload: dict[str, Any]) -> dict[str, Any]:
        """
        payload esperado:
        {
          "caminho_lote": "...xlsx",
          "coluna_processo": "Número do Processo",
          "evento_map": {
            "CALCP": "Agosto.2025 - PRINCIPAL",
            "HC30%": "Contratual - 30%",
            "HCP": "Contratual CHM",
            "CALCS": "Agosto.2025 - SUCUMBENCIA",
            "HSP": "Sucumb. Preço"
          },
          "solicitado_por": "45270",
          "modelo_path": null,
          "output_dir": null,
          "eventos": null
        }
        """
        result = gerar_lotes(
            caminho_lote=payload["caminho_lote"],
            coluna_processo=payload.get("coluna_processo", "Número do Processo"),
            evento_map=payload.get("evento_map", {}),
            solicitado_por=payload.get("solicitado_por", "45270"),
            modelo_path=payload.get("modelo_path"),
            output_dir=payload.get("output_dir"),
            eventos=payload.get("eventos"),
        )
        return asdict(result)
