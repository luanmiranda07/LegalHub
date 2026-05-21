from __future__ import annotations

import math
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Tuple

try:
    from pypdf import PdfReader, PdfWriter
except Exception as exc:  # pragma: no cover
    PdfReader = None  # type: ignore[assignment]
    PdfWriter = None  # type: ignore[assignment]
    PYPDF_IMPORT_ERROR = exc
else:
    PYPDF_IMPORT_ERROR = None


@dataclass
class SplitResult:
    output_files: List[Path]
    details: str


def _ensure_pypdf_available() -> None:
    if PYPDF_IMPORT_ERROR is not None:
        raise RuntimeError(
            "A biblioteca 'pypdf' nao esta instalada. Instale a dependencia do projeto e tente novamente."
        ) from PYPDF_IMPORT_ERROR


def sanitize_filename(name: str) -> str:
    name = str(name or "").strip()
    if not name:
        return "arquivo"
    name = re.sub(r'[\\/*?:"<>|]+', "_", name)
    name = re.sub(r"\s+", "_", name)
    return name[:150]


def ensure_output_dir(base_pdf: Path, output_dir: str | None) -> Path:
    if output_dir and str(output_dir).strip():
        out = Path(str(output_dir)).expanduser().resolve()
    else:
        out = base_pdf.with_name(f"{base_pdf.stem}_dividido")
    out.mkdir(parents=True, exist_ok=True)
    return out


def write_pdf(reader: PdfReader, page_indices: List[int], output_path: Path) -> None:  # type: ignore[valid-type]
    _ensure_pypdf_available()
    writer = PdfWriter()
    for idx in page_indices:
        writer.add_page(reader.pages[idx])
    with output_path.open("wb") as f:
        writer.write(f)


def split_one_per_page(pdf_path: Path, output_dir: Path, base_name: str) -> SplitResult:
    _ensure_pypdf_available()
    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    width = max(3, len(str(total_pages)))
    files: List[Path] = []

    for i in range(total_pages):
        output = output_dir / f"{sanitize_filename(base_name)}_{i + 1:0{width}d}.pdf"
        write_pdf(reader, [i], output)
        files.append(output)

    return SplitResult(files, f"PDF dividido em {total_pages} arquivos, um por pagina.")


def split_every_n_pages(pdf_path: Path, output_dir: Path, base_name: str, pages_per_file: int) -> SplitResult:
    _ensure_pypdf_available()
    if pages_per_file <= 0:
        raise ValueError("A quantidade de paginas por arquivo deve ser maior que zero.")

    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    total_files = math.ceil(total_pages / pages_per_file)
    width = max(3, len(str(total_files)))
    files: List[Path] = []

    file_no = 1
    for start in range(0, total_pages, pages_per_file):
        end = min(start + pages_per_file, total_pages)
        output = output_dir / f"{sanitize_filename(base_name)}_{file_no:0{width}d}_pag_{start + 1}_a_{end}.pdf"
        write_pdf(reader, list(range(start, end)), output)
        files.append(output)
        file_no += 1

    return SplitResult(files, f"PDF dividido em {len(files)} arquivos com ate {pages_per_file} paginas cada.")


def parse_ranges(ranges_text: str, total_pages: int) -> List[Tuple[int, int]]:
    if not str(ranges_text or "").strip():
        raise ValueError("Informe os intervalos de paginas.")

    ranges: List[Tuple[int, int]] = []
    parts = [p.strip() for p in str(ranges_text).split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            a_text, b_text = [x.strip() for x in part.split("-", 1)]
            if not a_text.isdigit() or not b_text.isdigit():
                raise ValueError(f"Intervalo invalido: {part}")
            a, b = int(a_text), int(b_text)
        else:
            if not part.isdigit():
                raise ValueError(f"Pagina invalida: {part}")
            a = b = int(part)

        if a < 1 or b < 1:
            raise ValueError(f"Paginas devem comecar em 1: {part}")
        if a > b:
            raise ValueError(f"Intervalo invertido: {part}")
        if b > total_pages:
            raise ValueError(f"O intervalo {part} ultrapassa o total de paginas do PDF ({total_pages}).")

        ranges.append((a - 1, b - 1))

    return ranges


def split_custom_ranges(pdf_path: Path, output_dir: Path, base_name: str, ranges_text: str) -> SplitResult:
    _ensure_pypdf_available()
    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    ranges = parse_ranges(ranges_text, total_pages)
    width = max(3, len(str(len(ranges))))
    files: List[Path] = []

    for idx, (start, end) in enumerate(ranges, start=1):
        output = output_dir / f"{sanitize_filename(base_name)}_{idx:0{width}d}_pag_{start + 1}_a_{end + 1}.pdf"
        write_pdf(reader, list(range(start, end + 1)), output)
        files.append(output)

    return SplitResult(files, f"PDF dividido em {len(files)} arquivos pelos intervalos informados.")


def extract_page_text(page: Any) -> str:
    try:
        return page.extract_text() or ""
    except Exception:
        return ""


def split_by_keyword(pdf_path: Path, output_dir: Path, base_name: str, keyword: str) -> SplitResult:
    _ensure_pypdf_available()
    keyword = str(keyword or "").strip()
    if not keyword:
        raise ValueError("Informe a palavra-chave.")

    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    normalized_keyword = keyword.casefold()

    hit_pages: List[int] = []
    for i, page in enumerate(reader.pages):
        text = extract_page_text(page)
        if normalized_keyword in text.casefold():
            hit_pages.append(i)

    if not hit_pages:
        raise ValueError(
            "Nenhuma pagina com a palavra-chave foi encontrada. "
            "Esse modo funciona melhor em PDFs de texto; PDFs escaneados podem exigir OCR."
        )

    files: List[Path] = []
    chunks: List[Tuple[int, int, str]] = []

    if hit_pages[0] > 0:
        chunks.append((0, hit_pages[0] - 1, "antes_da_palavra_chave"))

    for idx, start in enumerate(hit_pages):
        end = hit_pages[idx + 1] - 1 if idx + 1 < len(hit_pages) else total_pages - 1
        chunks.append((start, end, f"{idx + 1:03d}"))

    for start, end, suffix in chunks:
        output = output_dir / f"{sanitize_filename(base_name)}_{suffix}_pag_{start + 1}_a_{end + 1}.pdf"
        write_pdf(reader, list(range(start, end + 1)), output)
        files.append(output)

    details = (
        f"PDF dividido em {len(files)} arquivos com base na palavra-chave '{keyword}'.\n"
        f"Paginas com ocorrencia: {', '.join(str(i + 1) for i in hit_pages)}"
    )
    return SplitResult(files, details)


def _path_obrigatorio(valor: Any, nome: str) -> Path:
    texto = str(valor or "").strip()
    if not texto:
        raise ValueError(f"{nome} nao informado.")
    return Path(texto).expanduser().resolve()


def _validar_pdf(valor: Any) -> Path:
    pdf_path = _path_obrigatorio(valor, "PDF de entrada")
    if not pdf_path.exists() or not pdf_path.is_file():
        raise FileNotFoundError(f"PDF nao encontrado: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("O arquivo de entrada precisa estar no formato PDF.")
    return pdf_path


def _validar_paginas_por_arquivo(valor: Any) -> int:
    try:
        paginas = int(valor)
    except (TypeError, ValueError) as exc:
        raise ValueError("Informe uma quantidade de paginas valida.") from exc
    if paginas < 1:
        raise ValueError("A quantidade de paginas por arquivo deve ser maior ou igual a 1.")
    return paginas


def processar_payload(payload: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(payload, dict):
        payload = {}

    pdf_path = _validar_pdf(payload.get("pdfEntrada"))
    output_dir = ensure_output_dir(pdf_path, str(payload.get("pastaSaida") or ""))
    base_name = sanitize_filename(str(payload.get("nomeBase") or "").strip() or pdf_path.stem)
    mode = str(payload.get("modoDivisao") or "pagina").strip()

    if mode == "pagina":
        result = split_one_per_page(pdf_path, output_dir, base_name)
    elif mode == "quantidade":
        result = split_every_n_pages(
            pdf_path,
            output_dir,
            base_name,
            _validar_paginas_por_arquivo(payload.get("paginasPorArquivo")),
        )
    elif mode == "intervalos":
        result = split_custom_ranges(pdf_path, output_dir, base_name, str(payload.get("intervalos") or ""))
    elif mode == "palavra":
        result = split_by_keyword(pdf_path, output_dir, base_name, str(payload.get("palavraChave") or ""))
    else:
        raise ValueError("Modo de divisao invalido.")

    return {
        "details": result.details,
        "output_dir": str(output_dir),
        "output_files": [str(path) for path in result.output_files],
        "count": len(result.output_files),
    }


def abrir_pasta(caminho: Any) -> dict[str, Any]:
    pasta = _path_obrigatorio(caminho, "Pasta")
    if not pasta.exists() or not pasta.is_dir():
        raise FileNotFoundError(f"Pasta nao encontrada: {pasta}")

    if os.name == "nt":
        os.startfile(str(pasta))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(pasta)])
    else:
        subprocess.Popen(["xdg-open", str(pasta)])

    return {"path": str(pasta), "mensagem": "Pasta aberta."}
