from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple


PREFIX_DELIMITER_RE = re.compile(r"\.\s+")
NUMERIC_PREFIX_BEFORE_LAST_DOT_RE = re.compile(
    r"^\s*(?P<prefix>\d+(?:\s*[.-]\s*\d+)*)(?P<tail>\.(?:\s+.*)?)$"
)


def _suffix_after_prefix(name: str) -> str:
    match = PREFIX_DELIMITER_RE.search(name)
    if match:
        return name[match.end():].strip()
    if "." in name:
        return name.split(".", 1)[1].strip()
    return name


def build_plan_ccs_parceiro_sequencial(
    folder: Path,
    ccs_start: int | None,
    parceiro_start: int | None,
) -> List[Tuple[Path, str]]:
    """
    Regra:
      - Tudo antes do primeiro '.' e descartado
      - O restante (depois do '.') e mantido
      - Prefixo novo:
          * se CCS e Parceiro preenchidos -> "CCS.PARCEIRO."
          * se so CCS -> "CCS."
          * se so Parceiro -> "PARCEIRO."
      - CCS e/ou Parceiro incrementam +1 por pasta (somente os que existem)
    """
    plan: List[Tuple[Path, str]] = []

    ccs = ccs_start
    parceiro = parceiro_start

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        name = p.name.strip()
        suffix = _suffix_after_prefix(name)

        if ccs is not None and parceiro is not None:
            prefix = f"{ccs}.{parceiro}"
        elif ccs is not None:
            prefix = f"{ccs}"
        else:
            prefix = f"{parceiro}"

        new_name = f"{prefix}." if not suffix else f"{prefix}. {suffix}"
        plan.append((p, new_name))

        if ccs is not None:
            ccs += 1
        if parceiro is not None:
            parceiro += 1

    return plan


def build_plan_somente_css(
    folder: Path,
    ccs_start: int | None,
) -> List[Tuple[Path, str]]:
    plan: List[Tuple[Path, str]] = []
    ccs = ccs_start

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        name = p.name.strip()

        if ccs is not None:
            new_name = f"{ccs}. {name}"
            ccs += 1
        else:
            new_name = name

        plan.append((p, new_name))

    return plan


def build_plan_limpar_numeracao(folder: Path) -> List[Tuple[Path, str]]:
    """
    Remove a numeracao inicial ate o ultimo ponto do prefixo numerico,
    removendo tambem esse ponto do nome final.

    Ex.:
      "1. 7.7. Beltor - JUDICIAL - ..." -> "Beltor - JUDICIAL - ..."
      "7.7.1. Beltor - JUDICIAL - ..." -> "Beltor - JUDICIAL - ..."
      "104-123. DS Beline - ..." -> "DS Beline - ..."
    """
    plan: List[Tuple[Path, str]] = []

    subfolders = sorted(
        [x for x in folder.iterdir() if x.is_dir()],
        key=lambda x: x.name.lower(),
    )

    for p in subfolders:
        name = p.name.strip()
        match = NUMERIC_PREFIX_BEFORE_LAST_DOT_RE.match(name)

        if not match:
            plan.append((p, p.name))
            continue

        tail = match.group("tail").strip()
        new_name = tail[1:].strip() if tail.startswith(".") else tail
        if not new_name:
            plan.append((p, p.name))
            continue

        plan.append((p, new_name))

    return plan
