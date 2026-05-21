from __future__ import annotations

import os
import re
from collections import Counter
from pathlib import Path
from typing import Any

from openpyxl import Workbook


def normalizar_espacos(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip()


def limpar_numero(nome: str) -> str | None:
    padrao = r"^(?:CASO\s+)?\d+(?:\.\d+)?\s*-\s*"

    if re.match(padrao, nome, re.IGNORECASE):
        resto = re.sub(padrao, "", nome, flags=re.IGNORECASE).strip()
        return f"CASO - {resto}" if resto else None

    return None


def remover_prefixo_existente(nome_atual: str) -> str:
    nome = normalizar_espacos(nome_atual)

    nome = re.sub(
        r"^(?:CASO\s+)?\d+(?:\.\d+)?\s*-\s*",
        "",
        nome,
        flags=re.IGNORECASE,
    ).strip()

    nome = re.sub(
        r"^CASO\s*-\s*",
        "",
        nome,
        flags=re.IGNORECASE,
    ).strip()

    return nome


def montar_novo_nome(nome_atual: str, codigo: str, usar_so_numeracao: bool = False) -> str | None:
    resto = remover_prefixo_existente(nome_atual)

    if not resto:
        return None

    if usar_so_numeracao:
        return f"{codigo} - {resto}"

    return f"CASO {codigo} - {resto}"


def remover_sufixos_copia(texto: str) -> str:
    padrao = re.compile(
        r"(?:\s*-\s*c[óo]pia(?:\s*\(\d+\))?\s*)+$",
        re.IGNORECASE,
    )

    anterior = texto
    while True:
        novo = re.sub(padrao, "", anterior).strip()
        if novo == anterior:
            break
        anterior = novo

    return anterior


def extrair_nome_e_processo(nome_pasta: str) -> tuple[str, str] | None:
    nome_limpo = normalizar_espacos(nome_pasta)
    base = remover_prefixo_existente(nome_limpo)

    if not base:
        return None

    match = re.match(
        r"^(?P<nome>.*?)\s*-\s*Proc\.?\s*(?P<processo>.+?)\s*$",
        base,
        flags=re.IGNORECASE,
    )

    if not match:
        return None

    nome = normalizar_espacos(match.group("nome")).upper()
    processo = normalizar_espacos(match.group("processo"))
    processo = remover_sufixos_copia(processo)

    if not nome or not processo:
        return None

    return nome, processo


def _obter_pasta_base(pasta: Any) -> Path:
    pasta_base = str(pasta or "").strip()

    if not pasta_base:
        raise ValueError("Selecione uma pasta.")

    path_base = Path(pasta_base).expanduser().resolve()

    if not path_base.exists() or not path_base.is_dir():
        raise ValueError("A pasta selecionada é inválida.")

    return path_base


def _listar_pastas(path_base: Path) -> list[Path]:
    return sorted(
        [p for p in path_base.iterdir() if p.is_dir()],
        key=lambda p: p.name.lower(),
    )


def _payload_valor(payload: dict[str, Any], *nomes: str) -> Any:
    for nome in nomes:
        if nome in payload:
            return payload.get(nome)
    return None


def executar(payload: dict[str, Any] | None) -> dict[str, Any]:
    try:
        payload = payload or {}
        path_base = _obter_pasta_base(_payload_valor(payload, "pasta", "folderPath", "path"))

        numero_inicial = normalizar_espacos(_payload_valor(payload, "numeroInicial", "numero_inicial"))
        ano = normalizar_espacos(_payload_valor(payload, "ano"))
        usar_so_numeracao = bool(_payload_valor(payload, "usarSoNumeracao", "usar_so_numeracao"))

        pastas = _listar_pastas(path_base)
        if not pastas:
            return {
                "ok": False,
                "mensagem": "Nenhuma pasta encontrada.",
                "alteradas": 0,
                "ignoradas": [],
                "erros": [],
            }

        if numero_inicial == "" and ano == "":
            return _limpar_numeracao(path_base, pastas)

        if not numero_inicial.isdigit():
            return {
                "ok": False,
                "mensagem": "Informe um número inicial válido.",
                "alteradas": 0,
                "ignoradas": [],
                "erros": [],
            }

        if ano and not ano.isdigit():
            return {
                "ok": False,
                "mensagem": "Se informar o ano, ele deve ser numérico.",
                "alteradas": 0,
                "ignoradas": [],
                "erros": [],
            }

        return _adicionar_numeracao(
            path_base=path_base,
            pastas=pastas,
            numero_inicial=int(numero_inicial),
            ano=ano,
            usar_so_numeracao=usar_so_numeracao,
        )
    except Exception as exc:
        return {"ok": False, "mensagem": str(exc), "alteradas": 0, "ignoradas": [], "erros": [str(exc)]}


def _limpar_numeracao(path_base: Path, pastas: list[Path]) -> dict[str, Any]:
    alteradas = 0
    ignoradas: list[dict[str, str]] = []
    erros: list[dict[str, str]] = []

    for pasta in pastas:
        novo_nome = limpar_numero(pasta.name)

        if not novo_nome:
            ignoradas.append({"pasta": pasta.name, "motivo": "Sem numeração para remover."})
            continue

        novo_caminho = pasta.parent / novo_nome

        if novo_caminho.exists() and novo_caminho != pasta:
            ignoradas.append({"pasta": pasta.name, "motivo": f"Destino já existe: {novo_nome}"})
            continue

        try:
            os.rename(pasta, novo_caminho)
            alteradas += 1
        except OSError as exc:
            erros.append({"pasta": pasta.name, "motivo": str(exc)})

    return {
        "ok": not erros,
        "mensagem": f"{alteradas} pastas tiveram a numeração removida.",
        "pasta": str(path_base),
        "alteradas": alteradas,
        "ignoradas": ignoradas,
        "erros": erros,
    }


def _adicionar_numeracao(
    path_base: Path,
    pastas: list[Path],
    numero_inicial: int,
    ano: str,
    usar_so_numeracao: bool,
) -> dict[str, Any]:
    numero = numero_inicial
    alteradas = 0
    ignoradas: list[dict[str, str]] = []
    erros: list[dict[str, str]] = []

    for pasta in pastas:
        codigo = f"{numero}.{ano[-2:]}" if ano else str(numero)
        novo_nome = montar_novo_nome(pasta.name, codigo, usar_so_numeracao)

        if not novo_nome:
            ignoradas.append({"pasta": pasta.name, "motivo": "Nome sem conteúdo após remover prefixo."})
            continue

        novo_caminho = pasta.parent / novo_nome

        if novo_caminho.exists() and novo_caminho != pasta:
            ignoradas.append({"pasta": pasta.name, "motivo": f"Destino já existe: {novo_nome}"})
            continue

        try:
            os.rename(pasta, novo_caminho)
            numero += 1
            alteradas += 1
        except OSError as exc:
            erros.append({"pasta": pasta.name, "motivo": str(exc)})

    return {
        "ok": not erros,
        "mensagem": f"{alteradas} pastas renomeadas.",
        "pasta": str(path_base),
        "alteradas": alteradas,
        "ignoradas": ignoradas,
        "erros": erros,
    }


def identificar_duplicados(payload: dict[str, Any] | None) -> dict[str, Any]:
    try:
        payload = payload or {}
        path_base = _obter_pasta_base(_payload_valor(payload, "pasta", "folderPath", "path"))

        pastas = _listar_pastas(path_base)
        if not pastas:
            return {
                "ok": False,
                "mensagem": "Nenhuma pasta encontrada.",
                "arquivo_saida": "",
                "total_duplicados": 0,
                "ignoradas": [],
                "erros": [],
            }

        registros: list[tuple[str, str]] = []
        ignoradas: list[dict[str, str]] = []

        for pasta in pastas:
            dados = extrair_nome_e_processo(pasta.name)
            if dados is None:
                ignoradas.append({"pasta": pasta.name, "motivo": "Nome fora do padrão Nome - Proc. número."})
                continue
            registros.append(dados)

        if not registros:
            return {
                "ok": False,
                "mensagem": "Nenhuma pasta com padrão válido foi encontrada.",
                "arquivo_saida": "",
                "total_duplicados": 0,
                "ignoradas": ignoradas,
                "erros": [],
            }

        contagem = Counter(registros)
        duplicados = [
            {"nome": nome, "numero_processo": processo, "quantidade": quantidade}
            for (nome, processo), quantidade in contagem.items()
            if quantidade > 1
        ]

        if not duplicados:
            return {
                "ok": True,
                "mensagem": "Nenhuma duplicidade encontrada.",
                "arquivo_saida": "",
                "total_duplicados": 0,
                "ignoradas": ignoradas,
                "erros": [],
            }

        arquivo_saida = path_base / "duplicados_processos.xlsx"
        _salvar_duplicados(arquivo_saida, duplicados)

        return {
            "ok": True,
            "mensagem": f"{len(duplicados)} duplicidade(s) encontrada(s).",
            "arquivo_saida": str(arquivo_saida),
            "total_duplicados": len(duplicados),
            "duplicados": sorted(duplicados, key=lambda item: (item["nome"], item["numero_processo"])),
            "ignoradas": ignoradas,
            "erros": [],
        }
    except Exception as exc:
        return {
            "ok": False,
            "mensagem": str(exc),
            "arquivo_saida": "",
            "total_duplicados": 0,
            "ignoradas": [],
            "erros": [str(exc)],
        }


def _salvar_duplicados(arquivo_saida: Path, duplicados: list[dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Duplicados"

    ws.append(["nome", "numero_processo"])

    for item in sorted(duplicados, key=lambda row: (row["nome"], row["numero_processo"])):
        ws.append([item["nome"], item["numero_processo"]])

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 30

    wb.save(arquivo_saida)
