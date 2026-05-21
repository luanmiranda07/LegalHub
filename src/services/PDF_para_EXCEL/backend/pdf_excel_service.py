# src/services/PDF_para_EXCEL/backend/pdf_excel_service.py
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd


# =============================================================================
# Utilitários gerais
# =============================================================================

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()


def br_to_float(s: str) -> float:
    """Converte '12.345,67' -> 12345.67. Vazio/erro -> 0.0"""
    if not s:
        return 0.0
    s = str(s).strip()
    if not s:
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def float_to_br(x: float) -> str:
    """Converte 12345.67 -> '12.345,67' (PT-BR)."""
    try:
        s = f"{float(x):,.2f}"      # '12,345.67'
        s = s.replace(",", "X")     # '12X345.67'
        s = s.replace(".", ",")     # '12X345,67'
        s = s.replace("X", ".")     # '12.345,67'
        return s
    except Exception:
        return ""


# =============================================================================
# Extração - Seção 5 "Comparativo dos Valores"
# (mantida igual ao seu tratamento.py)
# =============================================================================

HEADER_PATTERN = r"5\.\s*Comparativo dos Valores"
NEXT_HEADER = r"\n\s*6\.\s*Observações|\Z"
NUM_PATTERN = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
DATE_PATTERN = r"\d{2}/\d{2}/\d{4}"


def extract_section5_from_text(full_text: str) -> str:
    """Pega o texto entre '5. Comparativo dos Valores' e '6. Observações'."""
    m = re.search(HEADER_PATTERN + r"(.*?)" + NEXT_HEADER, full_text, flags=re.S)
    return (m.group(1) if m else "").strip()


def normalize_lines(section_text: str) -> List[str]:
    """Quebra em linhas e compacta espaços."""
    return [
        re.sub(r"\s+", " ", ln).strip()
        for ln in section_text.splitlines()
        if ln.strip()
    ]


def parse_table(lines: List[str]) -> pd.DataFrame:
    """
    Constrói a tabela com colunas:
      Data | Valores (R$) | Principal (70%) | Contratual (30%) | Sucumbência | Total
    """
    rows: List[List[Any]] = []
    i = 0
    while i < len(lines):
        ln = lines[i]

        # pula cabeçalhos
        if ln.lower().startswith("data valores") or "(70%)" in ln or "(30%)" in ln:
            i += 1
            continue

        if re.match(DATE_PATTERN, ln):  # linha com data dd/mm/yyyy
            date = ln.split()[0]
            rest = ln[len(date):].strip()

            nums = re.findall(NUM_PATTERN, rest)

            first_num_match = re.search(NUM_PATTERN, rest)
            desc = rest[:first_num_match.start()].strip() if first_num_match else rest.strip()

            principal = contratual = sucumbencia = total = "0,00"
            if len(nums) >= 1:
                principal = nums[0]
            if len(nums) >= 2:
                contratual = nums[1]
            if len(nums) >= 3:
                sucumbencia = nums[2]
            if len(nums) >= 4:
                total = nums[3]

            rows.append([date, desc, principal, contratual, sucumbencia, total])
            i += 1
            continue

        if ln.startswith("Aguardando"):
            desc = ln
            j = i + 1
            while j < len(lines) and not (
                re.match(DATE_PATTERN, lines[j]) or lines[j].startswith("Aguardando")
            ):
                desc += " " + lines[j]
                j += 1
            desc_clean = desc.replace("Aguardando ", "", 1)
            rows.append(["Aguardando", desc_clean, "0,00", "0,00", "0,00", "0,00"])
            i = j
            continue

        # fallback
        rows.append([None, ln, "0,00", "0,00", "0,00", "0,00"])
        i += 1

    return pd.DataFrame(
        rows,
        columns=["Data", "Valores (R$)", "Principal (70%)", "Contratual (30%)", "Sucumbência", "Total"],
    )


def filter_most_recent(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna DF com apenas a linha da data mais recente (ignora 'Aguardando')."""
    valid_dates = df[df["Data"].str.match(DATE_PATTERN, na=False)].copy()
    if valid_dates.empty:
        return df.iloc[0:0].copy()

    valid_dates["Data_dt"] = pd.to_datetime(valid_dates["Data"], format="%d/%m/%Y", errors="coerce")
    idx = valid_dates["Data_dt"].idxmax()
    return df.loc[[idx]].reset_index(drop=True)


def processar_secao5_comparativo(texto: str) -> Tuple[str, str, str]:
    """Processa a seção 5 do PDF e retorna (contratual, sucumbencia, total=contratual+sucumb)."""
    section5_text = extract_section5_from_text(texto)
    if not section5_text:
        return "0,00", "0,00", "0,00"

    lines = normalize_lines(section5_text)
    df_sec5 = parse_table(lines)
    df_latest = filter_most_recent(df_sec5)

    if df_latest.empty:
        return "0,00", "0,00", "0,00"

    contratual_txt = df_latest.at[0, "Contratual (30%)"] or "0,00"
    sucumb_txt = df_latest.at[0, "Sucumbência"] or "0,00"

    total_val = br_to_float(contratual_txt) + br_to_float(sucumb_txt)
    return contratual_txt, sucumb_txt, float_to_br(total_val)


# =============================================================================
# Extração principal por PDF (campos que você já usa no projeto)
# =============================================================================

def extrair_dados_pdf(caminho_pdf: str) -> Dict[str, str]:
    """
    Extrai:
      - Nome (Parte Autora)
      - Número do processo (CNJ)
      - Data Encerramento
      - Contratual, Sucumbência, Total (da seção 5, linha mais recente)
    """
    caminho_pdf = str(caminho_pdf)
    p = Path(caminho_pdf)

    dados: Dict[str, str] = {
        "Arquivo": p.name,
        "Nome": "",
        "Número do processo": "",
        "Data encerramento": "",
        "CONTRATUAL": "",
        "Sucumbencia": "",
        "Total": "",
    }

    if not p.exists() or not p.is_file():
        raise FileNotFoundError(f"PDF não encontrado: {p}")

    with pdfplumber.open(str(p)) as pdf:
        # 1) extrai texto completo
        texto_pages: List[str] = []
        for page in pdf.pages:
            texto_pages.append(page.extract_text() or "")
        texto = "\n".join(texto_pages).strip()

        if not texto:
            # PDF escaneado/imagem geralmente cai aqui
            raise RuntimeError("PDF sem texto extraível (possível PDF escaneado/imagem).")

        # 2) Nome (Parte Autora) - regex
        m_aut = re.search(
            r"parte\s+autor(?:a|o)(?:\s*\(o\))?\s*[:\-–]?\s*([^\n\r|]+)",
            texto,
            flags=re.IGNORECASE,
        )
        if m_aut:
            dados["Nome"] = m_aut.group(1).strip()

        # fallback: procurar em tabelas (quando regex não acha)
        if not dados["Nome"]:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                achou = False
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        for i, cell in enumerate(row):
                            txt = str(cell) if cell else ""
                            if "parte autora" in _norm(txt):
                                nome = ""
                                if ":" in txt:
                                    nome = txt.split(":", 1)[1].strip()
                                if not nome and i + 1 < len(row):
                                    nome = (str(row[i + 1]) if row[i + 1] else "").strip()
                                if nome:
                                    dados["Nome"] = nome
                                    achou = True
                                    break
                        if achou:
                            break
                    if achou:
                        break
                if achou:
                    break

        # 3) Número do processo (CNJ)
        mproc = re.search(r"(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})", texto)
        if mproc:
            dados["Número do processo"] = mproc.group(1)

        # 4) Data de Encerramento
        m = re.search(
            r"(\d{2}[/-]\d{2}[/-]\d{4}).{0,60}Encerramento\s+com\s+a\s+Libera",
            texto,
            flags=re.IGNORECASE,
        )
        if m:
            dados["Data encerramento"] = m.group(1)
        else:
            m = re.search(
                r"Encerramento\s+com\s+a\s+Libera[^\n\r]{0,60}?(\d{2}[/-]\d{2}[/-]\d{4})",
                texto,
                flags=re.IGNORECASE,
            )
            if m:
                dados["Data encerramento"] = m.group(1)

        # 5) Seção 5 - Comparativo
        contratual, sucumbencia, total = processar_secao5_comparativo(texto)
        dados["CONTRATUAL"] = contratual
        dados["Sucumbencia"] = sucumbencia
        dados["Total"] = total

    return dados


def extrair_em_lote(pdf_paths: List[str]) -> List[Dict[str, Any]]:
    """
    Retorna uma lista de dicts (um por PDF), SEM quebrar o lote se um PDF falhar.
    Cada linha contém 'Erro' (string) quando houver problema.
    """
    resultados: List[Dict[str, Any]] = []

    for caminho in pdf_paths:
        try:
            d = extrair_dados_pdf(caminho)
            resultados.append(
                {
                    "Arquivo": d.get("Arquivo", Path(caminho).name),
                    "Nome": d.get("Nome", ""),
                    "Número do processo": d.get("Número do processo", ""),
                    "Data Encerramento": d.get("Data encerramento", ""),
                    "CONTRATUAL": d.get("CONTRATUAL", ""),
                    "Sucumbencia": d.get("Sucumbencia", ""),
                    "Total": d.get("Total", ""),
                    "Erro": "",
                }
            )
        except Exception as e:
            resultados.append(
                {
                    "Arquivo": Path(caminho).name,
                    "Nome": "",
                    "Número do processo": "",
                    "Data Encerramento": "",
                    "CONTRATUAL": "",
                    "Sucumbencia": "",
                    "Total": "",
                    "Erro": str(e),
                }
            )

    return resultados


# =============================================================================
# API pública do serviço (usada pelo WebViewApi)
# =============================================================================

def listar_pdfs_em_pasta(folder: str, *, recursive: bool = True) -> List[str]:
    """
    Lista PDFs em uma pasta (recursivo por padrão).
    Retorna lista de caminhos absolutos (strings), ordenados.
    """
    base = Path(str(folder)).expanduser().resolve()
    if not base.exists() or not base.is_dir():
        raise ValueError(f"Pasta inválida: {base}")

    it = base.rglob("*.pdf") if recursive else base.glob("*.pdf")
    files = [str(p.resolve()) for p in it if p.is_file()]

    # ordena por nome (case-insensitive)
    files.sort(key=lambda s: s.lower())
    return files


def processar_payload_dict(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Entrada (payload):
      {
        "pdf_paths": ["...pdf", ...] (opcional se folder_path existir)
        "folder_path": "..." (opcional)
        "recursive": true|false (opcional, default True)
      }

    Saída:
      {
        "files": [...],     # pdfs processados
        "rows": [...],      # linhas extraídas
        "summary": {...}    # contadores
      }
    """
    if payload is None:
        payload = {}

    recursive = bool(payload.get("recursive", True))

    pdf_paths_in = payload.get("pdf_paths")
    folder_path = payload.get("folder_path")

    files: List[str] = []

    if folder_path:
        files = listar_pdfs_em_pasta(str(folder_path), recursive=recursive)
    elif pdf_paths_in:
        if isinstance(pdf_paths_in, (str, bytes)):
            files = [str(pdf_paths_in)]
        else:
            files = [str(x) for x in (pdf_paths_in or [])]
    else:
        # nada informado
        return {
            "files": [],
            "rows": [],
            "summary": {"total": 0, "ok": 0, "with_data": 0, "errors": 0},
        }

    # Normaliza / resolve / remove duplicados mantendo ordem
    seen = set()
    files_norm: List[str] = []
    for f in files:
        p = Path(f).expanduser().resolve()
        s = str(p)
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        files_norm.append(s)

    rows = extrair_em_lote(files_norm)

    ok = sum(1 for r in rows if not r.get("Erro"))
    errors = sum(1 for r in rows if r.get("Erro"))
    with_data = sum(
        1
        for r in rows
        if (r.get("Número do processo") or r.get("Data Encerramento") or r.get("Total"))
        and not r.get("Erro")
    )

    return {
        "files": files_norm,
        "rows": rows,
        "summary": {"total": len(rows), "ok": ok, "with_data": with_data, "errors": errors},
    }


def gerar_excel_em_arquivo_lote(result_rows: List[Dict[str, Any]], caminho_saida: str) -> None:
    """
    Salva um Excel .xlsx a partir da lista de 'rows' retornada pelo processamento.
    """
    if not caminho_saida:
        raise ValueError("caminho_saida vazio")

    out = Path(str(caminho_saida)).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(result_rows or [])

    # Colunas padrão (mantendo seu padrão do tratamento.py)
    colunas_excel = [
        "Arquivo",
        "Nome",
        "Número do processo",
        "Data Encerramento",
        "CONTRATUAL",
        "Sucumbencia",
        "Total",
        "Erro",  # útil para auditoria; se você não quiser no excel, apaga essa linha
    ]

    for c in colunas_excel:
        if c not in df.columns:
            df[c] = ""

    # Se você NÃO quiser a coluna Erro no Excel, comente a linha acima e use:
    # colunas_excel = [c for c in colunas_excel if c != "Erro"]

    with pd.ExcelWriter(str(out), engine="openpyxl") as writer:
        df[colunas_excel].to_excel(writer, index=False, sheet_name="Dados")
