from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple


PREFIX_DELIMITER_RE = re.compile(r"\.\s+")
MULTISPACE_RE = re.compile(r"\s+")

STANDARD_JUDICIAL_RE = re.compile(
    r"\s-\sJUDICIAL\s-\s.+\s+x\s+INSS\b",
    flags=re.IGNORECASE,
)

NUMERIC_DASH_PREFIX_RE = re.compile(
    r"^\s*(?P<prefix>\d+(?:\s*-\s*\d+)+)(?P<tail>\.\s+.+)$"
)
NUMERIC_DOT_PREFIX_RE = re.compile(
    r"^\s*(?P<prefix>\d+(?:\s*\.\s*\d+)+)(?P<tail>\.\s+.+)$"
)


def _find_prefix_delimiter(name: str) -> int:
    match = PREFIX_DELIMITER_RE.search(name)
    if match:
        return match.start()
    return name.find(".")


def _norm_spaces(value: str) -> str:
    return MULTISPACE_RE.sub(" ", (value or "").strip())


def _replace_first_dash_prefix_with_dot(name: str) -> str:
    """
    Troca apenas o primeiro separador inicial por ponto.

    Regras:
      1. "- UNI ASSESSORIA - JUDICIAL - NOME x INSS - Proc"
         -> ". UNI ASSESSORIA - JUDICIAL - NOME x INSS - Proc"

      2. "549 - Beltor - JUDICIAL - NOME x INSS - MG - Proc"
         -> "549. Beltor - JUDICIAL - NOME x INSS - MG - Proc"

      3. "104-123. DS Beline - ..."
         -> "104.123. DS Beline - ..."
    """
    value = _norm_spaces(name)

    value = re.sub(r"^\s*-\s+", ". ", value, count=1)

    match = NUMERIC_DASH_PREFIX_RE.match(value)
    if match:
        prefix = re.sub(r"\s*-\s*", ".", match.group("prefix").strip())
        return f"{prefix}{match.group('tail')}"

    value = re.sub(r"^\s*(\d+)\s*-\s+", r"\1. ", value, count=1)
    return value


def _replace_first_dot_prefix_with_dash(name: str) -> str:
    """
    Troca ponto por hifen apenas no prefixo inicial.

    Regras:
      1. ". UNI ASSESSORIA - JUDICIAL - NOME x INSS - Proc"
         -> "- UNI ASSESSORIA - JUDICIAL - NOME x INSS - Proc"

      2. "549. Beltor - JUDICIAL - NOME x INSS - MG - Proc"
         -> "549 - Beltor - JUDICIAL - NOME x INSS - MG - Proc"

      3. "104.123. DS Beline - ..."
         -> "104-123. DS Beline - ..."
    """
    value = _norm_spaces(name)

    match = NUMERIC_DOT_PREFIX_RE.match(value)
    if match:
        prefix = re.sub(r"\s*\.\s*", "-", match.group("prefix").strip())
        return f"{prefix}{match.group('tail')}"

    value = re.sub(r"^\s*\.\s+", "- ", value, count=1)
    value = re.sub(r"^\s*(\d+)\s*\.\s+", r"\1 - ", value, count=1)
    return value


def _normalize_typed_prefix(prefix: str) -> str:
    """
    Normaliza prefixo digitado pelo usuario.

    Ex.:
      "104-123" -> "104.123"
      "104 - 123" -> "104.123"
    """
    prefix = (prefix or "").strip()

    if re.fullmatch(r"\d+(?:\s*-\s*\d+)+", prefix):
        return re.sub(r"\s*-\s*", ".", prefix)

    return prefix


def _fix_prefix_before_first_dot_name(name: str, new_prefix: str) -> str:
    prefix = _normalize_typed_prefix((new_prefix or "").strip())
    dot = _find_prefix_delimiter(name)

    if dot == -1:
        return name

    tail = name[dot:]

    if prefix == "":
        return tail.lstrip()

    if prefix == ".":
        return name

    return f"{prefix}{tail}"


def _uni_assessoria_judicial_name(name: str, partner_name: str) -> str:
    partner = _norm_spaces(partner_name)
    if not partner:
        raise ValueError("Informe o nome do parceiro.")

    value = _norm_spaces(name)
    if not value:
        return name

    value_fixed = _replace_first_dash_prefix_with_dot(value)

    if (
        value_fixed.upper().startswith(f". {partner.upper()} - JUDICIAL - ")
        or STANDARD_JUDICIAL_RE.search(value_fixed)
    ):
        return value_fixed

    return f". {partner} - JUDICIAL - {value_fixed} x INSS - Proc"


def build_plan_fix_prefix_before_first_dot(
    folder: Path,
    new_prefix: str,
) -> List[Tuple[Path, str]]:
    """
    Corrige o prefixo antes do primeiro "." no nome da pasta.

    Regras:
      - Remove tudo antes do primeiro "."
      - Mantem o "." e todo o restante intacto
      - Se nao existir ".", nao altera o nome
      - Se new_prefix for vazio, remove o prefixo e deixa apenas o tail
    """
    plan: List[Tuple[Path, str]] = []
    prefix = _normalize_typed_prefix((new_prefix or "").strip())

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        plan.append((p, _fix_prefix_before_first_dot_name(p.name, prefix)))

    return plan


def build_plan_uni_assessoria_judicial(
    folder: Path,
    partner_name: str = "UNI ASSESSORIA",
) -> List[Tuple[Path, str]]:
    """
    Padroniza nomes recebidos apenas como cliente.

    Ex.:
      "PEDRO SILVA DE MELO"
      -> ". UNI ASSESSORIA - JUDICIAL - PEDRO SILVA DE MELO x INSS - Proc"
    """
    plan: List[Tuple[Path, str]] = []
    partner = _norm_spaces(partner_name)

    if not partner:
        raise ValueError("Informe o nome do parceiro.")

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        plan.append((p, _uni_assessoria_judicial_name(p.name, partner)))

    return plan


def build_plan_replace_dash_with_dot_in_prefix(folder: Path) -> List[Tuple[Path, str]]:
    """
    Troca '-' por '.' apenas no prefixo inicial/numerico, mantendo o restante.
    """
    plan: List[Tuple[Path, str]] = []

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        plan.append((p, _replace_first_dash_prefix_with_dot(p.name)))

    return plan


def build_plan_replace_dot_with_dash_in_prefix(folder: Path) -> List[Tuple[Path, str]]:
    """
    Troca '.' por '-' apenas no prefixo inicial/numerico, mantendo o restante.
    """
    plan: List[Tuple[Path, str]] = []

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        plan.append((p, _replace_first_dot_prefix_with_dash(p.name)))

    return plan


def build_plan_corrigir_opcoes(
    folder: Path,
    new_prefix: str,
    *,
    padronizar_judicial: bool = False,
    partner_name: str = "UNI ASSESSORIA",
    trocar_hifen_por_ponto: bool = False,
    trocar_ponto_por_hifen: bool = False,
) -> List[Tuple[Path, str]]:
    """
    Monta um unico plano para as opcoes da tela de correcao.

    Quando nenhuma opcao especial esta marcada, aplica a correcao antiga do
    prefixo digitado. Quando alguma opcao especial esta marcada, o prefixo
    digitado fica preservado para evitar remocoes acidentais.
    """
    if trocar_hifen_por_ponto and trocar_ponto_por_hifen:
        raise ValueError("Escolha apenas uma conversao: '-' por '.' ou '.' por '-'.")

    plan: List[Tuple[Path, str]] = []
    has_special_mode = padronizar_judicial or trocar_hifen_por_ponto or trocar_ponto_por_hifen

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        new_name = p.name

        if padronizar_judicial:
            new_name = _uni_assessoria_judicial_name(new_name, partner_name)
        elif not has_special_mode:
            new_name = _fix_prefix_before_first_dot_name(new_name, new_prefix)

        if trocar_hifen_por_ponto:
            new_name = _replace_first_dash_prefix_with_dot(new_name)
        elif trocar_ponto_por_hifen:
            new_name = _replace_first_dot_prefix_with_dash(new_name)

        plan.append((p, new_name))

    return plan
