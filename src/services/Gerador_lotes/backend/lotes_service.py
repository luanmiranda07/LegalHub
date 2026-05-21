from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Any

import difflib
import re
import sys
import unicodedata

import numpy as np
import pandas as pd


DEFAULT_EVENTS = ["CALCP", "HC30%", "HCP", "CALCS", "HSP"]


def _src_base_dir() -> Path:
    """
    Base do app (src).
    - Em dev: .../src
    - Em PyInstaller: sys._MEIPASS aponta para a pasta extraída.
    """
    here = Path(__file__).resolve()
    # esperado: src/services/Gerador_lotes/backend/lotes_service.py
    src_dir = here.parents[3]  # backend -> Gerador_lotes -> services -> src
    return Path(getattr(sys, "_MEIPASS", src_dir))


def _default_model_path() -> Path:
    docs = _src_base_dir() / "docs"
    candidates = [
        docs / "modelo_lotes.xlsx",
        docs / "testesLotes.xlsx",
        docs / "modelo_excel_robo.xlsx",
        docs / "modelo_excel.xlsx",
    ]
    for p in candidates:
        if p.is_file():
            return p
    raise FileNotFoundError(
        "Modelo do Gerador de Lotes não encontrado.\n"
        "Coloque em src/docs/modelo_lotes.xlsx (recomendado)\n"
        "ou use src/docs/testesLotes.xlsx (compatibilidade)."
    )


def modelo_info() -> dict[str, str]:
    p = _default_model_path()
    return {"path": str(p), "name": p.name}


def _norm(s: Any) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s_nfkd = unicodedata.normalize("NFKD", s)
    s_noacc = "".join(c for c in s_nfkd if not unicodedata.combining(c))
    return s_noacc.lower()


def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols: list[str] = []
    for c in df.columns:
        cs = str(c).replace("\u00A0", " ")
        cs = re.sub(r"\s+", " ", cs).strip()
        new_cols.append(cs)
    df.columns = new_cols
    return df


def _find_best_column(df: pd.DataFrame, user_text: str | None) -> str | None:
    if not user_text:
        return None

    cols = list(df.columns)

    if user_text in cols:
        return user_text

    us = user_text.replace("\u00A0", " ").strip()
    for c in cols:
        if c.replace("\u00A0", " ").strip().lower() == us.lower():
            return c

    col_norm_map = {_norm(c): c for c in cols}
    usn = _norm(user_text)
    if usn in col_norm_map:
        return col_norm_map[usn]

    match = difflib.get_close_matches(usn, list(col_norm_map.keys()), n=1, cutoff=0.8)
    if match:
        return col_norm_map[match[0]]

    for kn, orig in col_norm_map.items():
        if usn in kn or kn in usn:
            return orig

    return None


def carregar_lote(caminho_lote: str, nome_coluna_processo: str | None = None) -> pd.DataFrame:
    raw = pd.read_excel(caminho_lote, header=None)

    alvo = (nome_coluna_processo.strip().lower() if nome_coluna_processo else "número do processo")
    header_row_idx = None

    # 1) tenta achar linha que contenha o texto da coluna de processo
    for i, row in raw.iterrows():
        linha_str = row.astype(str).str.strip().str.lower()
        if linha_str.eq(alvo).any():
            header_row_idx = i
            break

    # 2) fallback: primeira linha “cheia”
    if header_row_idx is None:
        for i, row in raw.iterrows():
            if row.count() >= 3:
                header_row_idx = i
                break

    if header_row_idx is None:
        raise ValueError("Não consegui localizar automaticamente a linha de cabeçalho no arquivo de lote.")

    header = raw.iloc[header_row_idx]
    dados = raw.iloc[header_row_idx + 1:].copy()
    dados.columns = header
    dados = dados.dropna(axis=1, how="all").dropna(how="all")
    dados = _normalize_headers(dados)
    return dados


def listar_colunas_lote(caminho_lote: str, coluna_processo_hint: str | None = None) -> list[str]:
    dados = carregar_lote(caminho_lote, nome_coluna_processo=coluna_processo_hint)
    return [str(c) for c in dados.columns]


def carregar_modelo(caminho_modelo: str) -> list[str]:
    modelo = pd.read_excel(caminho_modelo)
    return list(modelo.columns)


def montar_saida(
    dados_lote: pd.DataFrame,
    colunas_modelo: list[str],
    coluna_processo: str,
    evento_integracao_val: str,
    evento_map: dict[str, str],
    solicitado_por: str,
) -> tuple[pd.DataFrame, dict[str, str]]:
    saida = pd.DataFrame(columns=colunas_modelo)
    debug: dict[str, str] = {}

    col_proc = _find_best_column(dados_lote, coluna_processo)
    if col_proc:
        saida["PROCESSO"] = dados_lote[col_proc].values
        debug["PROCESSO"] = col_proc

    if evento_integracao_val in evento_map:
        coluna_origem_digitada = evento_map[evento_integracao_val]
        col_evt = _find_best_column(dados_lote, coluna_origem_digitada)
        if col_evt:
            saida["EVENTO"] = dados_lote[col_evt].values
            debug[f"EVENTO[{evento_integracao_val}]"] = col_evt

    if "DATA" in saida.columns:
        saida["DATA"] = date.today().strftime("%d/%m/%Y")
    if "RESULT" in saida.columns:
        saida["RESULT"] = "OK"
    if "SOLICITADO_POR" in saida.columns:
        saida["SOLICITADO_POR"] = solicitado_por
    if "EVENTO_INTEGRACAO" in saida.columns:
        saida["EVENTO_INTEGRACAO"] = evento_integracao_val

    return saida, debug


@dataclass
class GerarLotesResult:
    output_dir: str
    files: list[str]
    warnings: list[str]
    debug: dict[str, Any]


def gerar_lotes(
    caminho_lote: str,
    coluna_processo: str,
    evento_map: dict[str, str],
    solicitado_por: str = "45270",
    modelo_path: str | None = None,
    output_dir: str | None = None,
    eventos: list[str] | None = None,
) -> GerarLotesResult:
    lote_path = Path(caminho_lote)
    if not lote_path.is_file():
        raise FileNotFoundError(f"Arquivo de lote não encontrado: {caminho_lote}")

    model_path = Path(modelo_path) if modelo_path else _default_model_path()
    if not model_path.is_file():
        raise FileNotFoundError(f"Modelo não encontrado: {str(model_path)}")

    out_dir = Path(output_dir) if output_dir else lote_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    dados_lote = carregar_lote(str(lote_path), nome_coluna_processo=coluna_processo or None)
    colunas_modelo = carregar_modelo(str(model_path))

    evs = eventos or DEFAULT_EVENTS

    generated: list[str] = []
    warnings: list[str] = []
    debug: dict[str, Any] = {"model": str(model_path), "mappings": {}, "events": evs}

    consolidados: list[pd.DataFrame] = []

    for idx, evento in enumerate(evs, start=1):
        saida, dbg_map = montar_saida(
            dados_lote=dados_lote,
            colunas_modelo=colunas_modelo,
            coluna_processo=coluna_processo,
            evento_integracao_val=evento,
            evento_map=evento_map,
            solicitado_por=solicitado_por,
        )
        debug["mappings"][evento] = dbg_map

        if "PROCESSO" not in dbg_map:
            warnings.append(f"{evento}: coluna de PROCESSO não encontrada pelo texto '{coluna_processo}'.")
        if f"EVENTO[{evento}]" not in dbg_map:
            warnings.append(f"{evento}: coluna do valor não encontrada pelo texto '{evento_map.get(evento, '')}'.")

        consolidados.append(saida)

        out_file = out_dir / f"{idx} Cópia de modelo rb 03 - {evento} preenchido.xlsx"
        saida.to_excel(out_file, index=False)
        generated.append(str(out_file))

    saida_todos = pd.concat(consolidados, ignore_index=True)
    out_all = out_dir / "6 Cópia de modelo rb 03 - TODOS preenchido.xlsx"
    saida_todos.to_excel(out_all, index=False)
    generated.append(str(out_all))

    return GerarLotesResult(
        output_dir=str(out_dir),
        files=generated,
        warnings=warnings,
        debug=debug,
    )


def gerar_lotes_dict(payload: dict[str, Any]) -> dict[str, Any]:
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
