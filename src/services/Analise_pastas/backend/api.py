from __future__ import annotations

import traceback
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, List

import webview
import openpyxl

from .preencher_pastas import (
    listar_subpastas,
    parse_nome_pasta,
    validar,
    validar_robo,
    mapear_colunas_por_header,
    detectar_template,
    EXPECTED_HEADERS_PADRAO,
    EXPECTED_HEADERS_ROBO,
)

def _ok(data: Any = None) -> Dict[str, Any]:
    return {"ok": True, "data": data}

def _err(msg: str, details: str | None = None) -> Dict[str, Any]:
    out: Dict[str, Any] = {"ok": False, "error": msg}
    if details:
        out["details"] = details
    return out

class Api:
    def __init__(self) -> None:
        self.window: webview.Window | None = None

    def dialogo_pasta(self) -> Dict[str, Any]:
        try:
            if not self.window:
                return _err("Window não inicializada.")
            paths = self.window.create_file_dialog(webview.FOLDER_DIALOG, allow_multiple=False)
            return _ok(paths[0] if paths else None)
        except Exception:
            return _err("Falha ao abrir diálogo de pasta.", traceback.format_exc())

    def dialogo_arquivo_excel(self) -> Dict[str, Any]:
        try:
            if not self.window:
                return _err("Window não inicializada.")
            file_types = ("Excel (*.xlsx;*.xlsm)", "*.xlsx;*.xlsm")
            paths = self.window.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=[file_types],
            )
            return _ok(paths[0] if paths else None)
        except Exception:
            return _err("Falha ao abrir diálogo de arquivo.", traceback.format_exc())

    def analisar_pastas(self, pasta_raiz: str, template_excel: str | None = None) -> Dict[str, Any]:
        """
        Retorna parse + validação das subpastas.
        Se template_excel for informado, detecta PADRAO vs ROBO por colunas.
        """
        try:
            root = Path(pasta_raiz)
            if not root.is_dir():
                return _err(f"Pasta inválida: {pasta_raiz}")

            kind = "PADRAO"
            faltantes: List[str] = []

            if template_excel:
                tpl = Path(template_excel)
                if not tpl.is_file():
                    return _err(f"Template inválido: {template_excel}")

                wb = openpyxl.load_workbook(tpl)
                ws = wb.active
                col_map = mapear_colunas_por_header(ws)
                kind = detectar_template(col_map)

                expected = EXPECTED_HEADERS_ROBO if kind == "ROBO" else EXPECTED_HEADERS_PADRAO
                faltantes = [h for h in expected if h.upper() not in col_map]

            subpastas = listar_subpastas(root)
            itens: List[Dict[str, Any]] = []

            ok_count = 0
            problema_count = 0

            for p in subpastas:
                reg = parse_nome_pasta(p.name)
                erros = validar_robo(reg) if kind == "ROBO" else validar(reg)

                # considera "OK" se não houver erro "de verdade" (avisos não contam)
                is_ok = len([e for e in erros if not e.lower().startswith("aviso:")]) == 0

                item = asdict(reg)
                item.update(
                    {
                        "path": str(p),
                        "erros": erros,
                        "is_ok": is_ok,
                    }
                )
                itens.append(item)

                if is_ok:
                    ok_count += 1
                else:
                    problema_count += 1

            return _ok(
                {
                    "template_detectado": kind,
                    "faltantes_template": faltantes,
                    "total": len(itens),
                    "ok": ok_count,
                    "com_problemas": problema_count,
                    "itens": itens,
                }
            )

        except Exception:
            return _err("Erro inesperado ao analisar pastas.", traceback.format_exc())
