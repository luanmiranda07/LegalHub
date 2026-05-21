"""
webview_api.py (COMENTADO)

Este arquivo é a “ponte” entre o Frontend (JavaScript/HTML) e o Backend (Python)
quando você roda o app via pywebview.

No pywebview, tudo que você expõe no parâmetro `js_api=...` ao criar a janela
fica acessível no JS em:  window.pywebview.api.<metodo>()

Ou seja:
- JS chama:   window.pywebview.api.pick_folder()
- Python executa aqui: WebViewApi.pick_folder()
- Python retorna dict/list/str/int/bool (serializável)
- JS recebe o retorno e atualiza a UI.

A ideia principal:
1) abrir diálogos (pasta/arquivo) no desktop
2) chamar as rotinas do serviço Analise_pastas
3) devolver respostas “amigáveis” para o JS (logs/ok/error/etc.)
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from uuid import uuid4

import webview
from backend.engine import LoginEngine

# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: Analise_pastas
# -----------------------------------------------------------------------------
# Esses imports são as “funções motor” do seu backend real.
# A WebViewApi só orquestra chamada + validação + retorno serializável.

# DS Beline
from services.Analise_pastas.backend.preencher_excel_parceiro import preencher_excel_parceiro
from services.Analise_pastas.backend.preencher_codigo_beline import build_plan_codigo_beline

# botão "Executar Somente CCS"
from services.Analise_pastas.backend.renomear_pastas import (
    build_plan_ccs_parceiro_sequencial,
    build_plan_limpar_numeracao,
    build_plan_somente_css,
)

from services.Analise_pastas.backend.preencher_pastas import preencher_excel
from services.Analise_pastas.backend.corrigir_numeracao import (
    build_plan_corrigir_opcoes,
)


# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: Gerador_lotes
# -----------------------------------------------------------------------------
# Este serviço gera arquivos Excel de integração (1 por evento + 1 consolidado).
# O frontend chama via pywebview e o backend (pandas/openpyxl) grava no disco.
from services.Gerador_lotes.backend.lotes_service import (
    modelo_info as lotes_modelo_info,
    listar_colunas_lote,
    gerar_lotes_dict,
)

# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: PDF_para_EXCEL
# -----------------------------------------------------------------------------
# Serviço para extrair dados de PDFs (em lote) e gerar Excel.
from services.PDF_para_EXCEL.backend.pdf_excel_service import (
    listar_pdfs_em_pasta,
    processar_payload_dict as pdfexcel_process_payload_dict,
    gerar_excel_em_arquivo_lote,
)

from services.divisor_pdf.backend.divisorpdf_servico import (
    abrir_pasta as divisorpdf_abrir_pasta_service,
    processar_payload as divisorpdf_processar_payload,
)


# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: IA_LAUDO
# -----------------------------------------------------------------------------

from services.IA_LAUDO.backend.ialaudos import analisar_pdf as ialaudos_analisar_pdf

# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: IA_PET
# -----------------------------------------------------------------------------
from services.IA_PET_INICIAL.backend.iapet import (
    analisar_pdfs as iapet_analisar_pdfs,
    gerar_word_de_resposta as iapet_gerar_word_de_resposta,
)



# -----------------------------------------------------------------------------
# IMPORTS DO SERVIÇO: RELATORIO CONFORMIDADE
# -----------------------------------------------------------------------------

from services.Relatorio_Conformidade.backend.relatorio_service import gerar_relatorios as relconf_gerar_relatorios

# -----------------------------------------------------------------------------
# IMPORTS DO SERVICO: Gerador de Documentos
# -----------------------------------------------------------------------------
from services.gerador_documentos.backend.geradordocumentos_service import (
    abrir_pasta as gerdoc_abrir_pasta_service,
    cancelar_job as gerdoc_cancelar_job,
    iniciar_job as gerdoc_iniciar_job,
    modelo_info as gerdoc_modelo_info_service,
    obter_status_job as gerdoc_obter_status_job,
    validar_payload as gerdoc_validar_payload,
)

# -----------------------------------------------------------------------------
# IMPORTS DO SERVICO: AnalistDuplic
# -----------------------------------------------------------------------------
from services.AnalistDuplic.backend.analisduplic_serveco import (
    executar as analisduplic_executar_service,
    identificar_duplicados as analisduplic_identificar_duplicados_service,
)

# -----------------------------------------------------------------------------
# IMPORTS DO SERVICO: ExitoBot
# -----------------------------------------------------------------------------
from services.ExitoBot.backend.exitobot_servico import processar_payload as exitobot_processar_payload

# -----------------------------------------------------------------------------
# Regras de nomes no Windows
# -----------------------------------------------------------------------------
# No Windows, pastas não podem conter alguns caracteres e alguns nomes são
# reservados pelo sistema (CON, PRN, AUX, NUL, COM1..COM9, LPT1..LPT9).
#
# Como seu app renomeia pastas, isso evita crash/erro durante rename.
_INVALID_WIN_CHARS_RE = re.compile(r'[<>:"/\\\\|?*]')

_RESERVED_WIN_NAMES = {
    "con", "prn", "aux", "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}


def _is_windows() -> bool:
    """Detecta se o sistema operacional atual é Windows."""
    return os.name == "nt"


def _validate_folder_name(name: str) -> Optional[str]:
    """
    Valida se um nome de pasta é aceitável (principalmente no Windows).

    Retorna:
      - None            => nome OK
      - str com erro    => motivo do nome ser inválido

    Observação:
      - Em Linux/Mac, as regras são mais permissivas.
      - Aqui a validação “forte” só é aplicada em Windows.
    """
    if name is None:
        return "Nome vazio (None)."

    n = str(name)

    # Nome só com espaços conta como vazio.
    if not n.strip():
        return "Nome vazio."

    if _is_windows():
        # Caracteres proibidos no Windows
        if _INVALID_WIN_CHARS_RE.search(n):
            return 'Nome contém caracteres inválidos para Windows: <>:"/\\|?*'

        # No Windows, não pode terminar com espaço ou ponto.
        if n.endswith(" ") or n.endswith("."):
            return "Nome não pode terminar com espaço ou ponto (Windows)."

        # O Windows trata alguns nomes como “dispositivos”.
        # Ex.: "CON", "AUX" etc. Mesmo com extensão, pode dar problema.
        base = n.split(".", 1)[0].strip().lower()
        if base in _RESERVED_WIN_NAMES:
            return f"Nome reservado no Windows: {base.upper()}"

    return None


def _count_subfolders(folder: Path) -> int:
    """
    Conta quantas subpastas DIRETAS existem dentro de `folder`.

    Ex.: se folder tem:
      - a/ (dir)
      - b/ (dir)
      - arquivo.txt (file)
    retorna 2
    """
    return sum(1 for p in folder.iterdir() if p.is_dir())


def _timestamp() -> str:
    """Gera timestamp no formato AAAAMMDD_HHMMSS para usar em nomes de arquivo."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _unique_temp_name(i: int) -> str:
    """
    Gera um nome temporário único para renomeações em 2 fases.

    Por quê?
      - Para evitar conflitos quando você renomeia várias pastas “em cadeia”.
      - Ex.: A -> B e B -> C pode colidir se feito em uma fase só.
    """
    return f"__tmp__{uuid4().hex}__{i}"


def _safe_rename_plan(pairs: List[Tuple[Path, str]], *, dry_run: bool) -> Dict[str, Any]:
    """
    Executa (ou apenas simula) um plano de renomeações.

    Entrada:
      pairs: lista de (old_path, new_name)
        - old_path é o Path da pasta existente
        - new_name é o *nome* final (apenas o nome, não caminho inteiro)

      dry_run:
        - True  => NÃO renomeia nada, só devolve logs de prévia
        - False => renomeia de verdade

    Saída (sempre dict serializável pro JS):
      {
        ok: bool,
        error?: str,
        logs: [str, ...],
        renamed: int,
        skipped: int,
        errors: int
      }

    Por que existe?
      - Centraliza validações (nome inválido, duplicata, conflito com pasta existente)
      - Faz a renomeação em 2 fases (old -> tmp -> final) para reduzir riscos
    """
    logs: List[str] = []

    # Aqui vamos transformar (old_path, new_name) em (old_path, new_path)
    # e filtrar o que deve ser pulado.
    items: List[Tuple[Path, Path]] = []
    skipped = 0
    errors = 0

    # -------------------------------------------------------------------------
    # 1) Validações básicas + montar lista de renomeações reais
    # -------------------------------------------------------------------------
    for old_path, new_name in pairs:
        if not old_path.exists():
            skipped += 1
            logs.append(f"[PULAR] Não existe: {old_path}")
            continue

        err = _validate_folder_name(new_name)
        if err:
            errors += 1
            logs.append(f"[ERRO] Nome inválido: '{new_name}' ({err}) • pasta={old_path.name}")
            continue

        new_path = old_path.parent / new_name

        # Se o nome final for igual ao atual, não há o que fazer.
        if old_path.name == new_name:
            skipped += 1
            logs.append(f"[PULAR] Sem mudança: {old_path.name}")
            continue

        items.append((old_path, new_path))

    # -------------------------------------------------------------------------
    # 2) Evitar “duplicatas” de destino (duas pastas tentando virar o mesmo nome)
    # -------------------------------------------------------------------------
    seen: Dict[str, Path] = {}
    for _, new_path in items:
        # No Windows, o filesystem costuma ser case-insensitive,
        # então "ABC" e "abc" são considerados o mesmo nome.
        key = new_path.name.lower() if _is_windows() else new_path.name
        if key in seen:
            errors += 1
            logs.append(f"[ERRO] Dois destinos iguais: '{new_path.name}' (conflito com '{seen[key].name}')")
        else:
            seen[key] = new_path

    if errors:
        # Se já tem erro aqui, devolve e não tenta renomear.
        return {
            "ok": False,
            "error": "Plano contém erros (nomes inválidos/duplicados).",
            "logs": logs,
            "renamed": 0,
            "skipped": skipped,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # 3) Evitar conflito com pastas que já existem NO DISCO e não fazem parte
    #    do conjunto “old” do plano.
    #
    # Ex.: você quer renomear X -> Y, mas Y já existe e não vai “sumir” no plano.
    # -------------------------------------------------------------------------
    old_set = {old.resolve() for old, _ in items}
    for _, new_path in items:
        if new_path.exists() and new_path.resolve() not in old_set:
            errors += 1
            logs.append(f"[ERRO] Destino já existe e não faz parte do plano: {new_path}")

    if errors:
        return {
            "ok": False,
            "error": "Conflito com pastas já existentes no destino.",
            "logs": logs,
            "renamed": 0,
            "skipped": skipped,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # 4) Modo prévia: apenas mostra o que faria, sem renomear.
    # -------------------------------------------------------------------------
    if dry_run:
        for old, new in items:
            logs.append(f"[PRÉVIA] {old.name}  ->  {new.name}")

        # Nota: aqui o código retorna renamed=0 e soma “skipped” com len(items)
        # para indicar “nada foi aplicado” (tudo ficou em prévia).
        return {"ok": True, "logs": logs, "renamed": 0, "skipped": skipped + len(items), "errors": 0}

    # -------------------------------------------------------------------------
    # 5) Execução real (2 fases)
    #
    # Fase A: renomeia tudo para temporários únicos (old -> tmp)
    # Fase B: renomeia temporários para o destino final (tmp -> final)
    #
    # Isso evita colisão quando o destino final “bate” em outro nome existente
    # que também será renomeado.
    # -------------------------------------------------------------------------
    temp_map: List[Tuple[Path, Path]] = []
    renamed = 0

    try:
        # Fase A
        for i, (old, final) in enumerate(items, start=1):
            tmp = old.parent / _unique_temp_name(i)
            old.rename(tmp)
            temp_map.append((tmp, final))
    except Exception as e:
        # Se falhar aqui, pelo menos não tentamos fase B.
        return {"ok": False, "error": str(e), "logs": logs, "renamed": renamed, "skipped": skipped, "errors": 1}

    try:
        # Fase B
        for tmp, final in temp_map:
            tmp.rename(final)
            renamed += 1
            logs.append(f"[OK] {final.name}")
    except Exception as e:
        # Se falhar na fase B, pode existir estado “meio aplicado”.
        # (um upgrade futuro seria tentar rollback)
        return {"ok": False, "error": str(e), "logs": logs, "renamed": renamed, "skipped": skipped, "errors": 1}

    return {"ok": True, "logs": logs, "renamed": renamed, "skipped": skipped, "errors": 0}


# -----------------------------------------------------------------------------
# CLASSE PRINCIPAL EXPOSTA AO FRONTEND (window.pywebview.api)
# -----------------------------------------------------------------------------
@dataclass
class WebViewApi:
    """
    Esta classe é o “objeto API” exposto ao JavaScript via pywebview.

    Quando no Python você cria a janela assim:
        api = WebViewApi(...)
        webview.create_window(..., js_api=api)

    então no JavaScript aparece:
        window.pywebview.api.pick_folder()
        window.pywebview.api.run_preencher({...})
        ...

    Métodos esperados pelo Analise_pastas/frontend/analise_de_pastas.js:
      - pick_folder()
      - pick_template()
      - run_preencher({folderPath, templateMode, templatePath})
      - run_renomear({folderPath, ccsStart, parceiroStart, dryRun, limparNumeracao})
      - run_corrigir({folderPath, prefix, dryRun, padronizarJudicial, trocarHifenPorPonto, trocarPontoPorHifen})
    """
    base_dir: Path
    window: webview.Window | None = None
    _login_engine: LoginEngine | None = field(default=None, init=False, repr=False)

    def attach_window(self, window: webview.Window) -> None:
        """
        Guarda a referência da janela.

        Por quê isso é útil?
          - Para abrir dialogs (create_file_dialog) você precisa da window.
          - Em alguns casos, pywebview também permite pegar webview.windows[0],
            mas manter a referência é mais confiável.
        """
        self.window = window

    def _require_window(self) -> webview.Window:
        """
        Garante que temos uma Window para usar nos diálogos.

        Estratégia:
          1) se `self.window` foi setada (attach_window), usa ela
          2) senão, tenta pegar a primeira janela global (webview.windows[0])
          3) se não existir, lança erro
        """
        if self.window is not None:
            return self.window
        if webview.windows:
            return webview.windows[0]
        raise RuntimeError("Janela pywebview não inicializada.")

    def _get_login_engine(self) -> LoginEngine:
        """Inicializa o engine de login sob demanda e reutiliza a instância."""
        if self._login_engine is not None:
            return self._login_engine

        host = (os.getenv("LDAP_HOST") or "").strip()
        domain_fqdn = (os.getenv("DOMAIN_FQDN") or "").strip()
        group_dn = (os.getenv("GROUP_DN") or "").strip() or None

        if not host:
            raise RuntimeError("Variavel LDAP_HOST nao configurada.")
        if not domain_fqdn:
            raise RuntimeError("Variavel DOMAIN_FQDN nao configurada.")

        self._login_engine = LoginEngine(
            host=host,
            domain_fqdn=domain_fqdn,
            group_dn=group_dn,
        )
        return self._login_engine

    def login(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Autentica no LDAP para o fluxo de login da janela pywebview."""
        try:
            if not isinstance(payload, dict):
                payload = {}

            usuario = str(payload.get("usuario") or "").strip()
            senha = str(payload.get("senha") or "").strip()

            if not usuario or not senha:
                return {
                    "ok": False,
                    "code": 400,
                    "erro": "Preencha usuario e senha.",
                }

            engine = self._get_login_engine()
            resp = engine.authenticate(usuario, senha)
            resp_dict = dict(resp or {})

            code = int(resp_dict.get("code", 500))
            resp_dict["code"] = code
            resp_dict["ok"] = code == 200

            if not resp_dict["ok"] and not resp_dict.get("erro"):
                resp_dict["erro"] = "Credenciais invalidas. Verifique usuario e senha."

            return resp_dict
        except Exception as e:
            return {"ok": False, "code": 500, "erro": str(e)}

    # -------------------------------------------------------------------------
    # Dialog: selecionar pasta base
    # -------------------------------------------------------------------------
    def pick_folder(self) -> Dict[str, Any]:
        """
        Abre um diálogo nativo para o usuário escolher uma pasta.

        Retorno para o JS:
          - {"canceled": True} se o usuário cancelou
          - ou info da pasta (folderPath, name, subfolderCount)
        """
        w = self._require_window()

        # webview.FOLDER_DIALOG abre “selecionar pasta”.
        paths = w.create_file_dialog(webview.FOLDER_DIALOG)

        if not paths:
            return {"canceled": True}

        folder_path = Path(paths[0]).resolve()

        # Segurança: garantir que é diretório existente.
        if not folder_path.exists() or not folder_path.is_dir():
            return {"canceled": False, "ok": False, "error": f"Diretório inválido: {folder_path}"}

        return {
            "canceled": False,
            "folderPath": str(folder_path),
            "name": folder_path.name,
            "subfolderCount": _count_subfolders(folder_path),
        }

    # -------------------------------------------------------------------------
    # ExitoBot
    # -------------------------------------------------------------------------
    def exitobot_pick_excel(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            paths = w.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("Planilhas (*.xlsx;*.xlsm;*.xls)",),
            )

            if not paths:
                return {"canceled": True}

            if isinstance(paths, (list, tuple)):
                path = paths[0] if paths else ""
            else:
                path = paths

            excel_path = Path(str(path)).resolve()

            if not excel_path.exists() or not excel_path.is_file():
                return {"canceled": False, "ok": False, "error": f"Arquivo Excel invalido: {excel_path}"}

            if excel_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
                return {"canceled": False, "ok": False, "error": "Selecione uma planilha .xlsx, .xlsm ou .xls."}

            return {
                "canceled": False,
                "ok": True,
                "path": str(excel_path),
                "excelPath": str(excel_path),
                "name": excel_path.name,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def exitobot_validar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            return exitobot_processar_payload(payload, dry_run=True)
        except Exception as e:
            return {"ok": False, "mensagem": str(e), "error": str(e), "logs": [f"[ERRO] {e}"]}

    def exitobot_executar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            return exitobot_processar_payload(payload, dry_run=False)
        except Exception as e:
            return {"ok": False, "mensagem": str(e), "error": str(e), "logs": [f"[ERRO] {e}"]}

    # -------------------------------------------------------------------------
    # AnalistDuplic
    # -------------------------------------------------------------------------
    def analisduplic_pick_folder(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            paths = w.create_file_dialog(webview.FOLDER_DIALOG)

            if not paths:
                return {"canceled": True}

            if isinstance(paths, (list, tuple)):
                p = paths[0] if paths else ""
            else:
                p = paths

            folder_path = Path(str(p)).resolve()

            if not folder_path.exists() or not folder_path.is_dir():
                return {"canceled": False, "ok": False, "error": f"Diretorio invalido: {folder_path}"}

            return {
                "canceled": False,
                "ok": True,
                "path": str(folder_path),
                "folderPath": str(folder_path),
                "name": folder_path.name,
                "subfolderCount": _count_subfolders(folder_path),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def analisduplic_executar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            return analisduplic_executar_service(payload)
        except Exception as e:
            return {"ok": False, "mensagem": str(e), "error": str(e)}

    def analisduplic_identificar_duplicados(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            return analisduplic_identificar_duplicados_service(payload)
        except Exception as e:
            return {"ok": False, "mensagem": str(e), "error": str(e)}

    # -------------------------------------------------------------------------
    # Dialog: selecionar template Excel
    # -------------------------------------------------------------------------
    def pick_template(self) -> Dict[str, Any]:
        """
        Abre um diálogo para o usuário escolher um arquivo .xlsx (template).

        Retorno:
          - {"canceled": True} se cancelar
          - {"canceled": False, "templatePath": "...", "name": "..."} se ok
        """
        w = self._require_window()

        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx)",),
        )

        if not paths:
            return {"canceled": True}

        p = Path(paths[0]).resolve()

        if not p.exists() or not p.is_file():
            return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {p}"}

        return {"canceled": False, "templatePath": str(p), "name": p.name}


    # -------------------------------------------------------------------------
    # Dialog: selecionar arquivo de LOTE (Excel)
    # -------------------------------------------------------------------------
    def pick_lote_file(self) -> Dict[str, Any]:
        """
        Abre um diálogo para o usuário escolher um arquivo .xlsx (LOTE).

        Retorno:
          - {"canceled": True} se cancelar
          - {"canceled": False, "lotePath": "...", "name": "..."} se ok
        """
        w = self._require_window()

        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx)",),
        )

        if not paths:
            return {"canceled": True}

        p = Path(paths[0]).resolve()
        if not p.exists() or not p.is_file():
            return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {p}"}

        return {"canceled": False, "lotePath": str(p), "name": p.name}

    # -------------------------------------------------------------------------
    # Dialog: selecionar modelo manual (Excel) — opcional
    # -------------------------------------------------------------------------
    def pick_model_file(self) -> Dict[str, Any]:
        """
        Abre um diálogo para o usuário escolher um arquivo .xlsx (modelo).

        Observação:
          - Se o usuário NÃO selecionar, o backend usa o modelo padrão em src/docs.

        Retorno:
          - {"canceled": True} se cancelar
          - {"canceled": False, "modelPath": "...", "name": "..."} se ok
        """
        w = self._require_window()

        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx)",),
        )

        if not paths:
            return {"canceled": True}

        p = Path(paths[0]).resolve()
        if not p.exists() or not p.is_file():
            return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {p}"}

        return {"canceled": False, "modelPath": str(p), "name": p.name}

    # -------------------------------------------------------------------------
    # AÇÃO: Preencher excel com base nas pastas
    # -------------------------------------------------------------------------
    def run_preencher(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            folder_path = Path(str(payload.get("folderPath") or "")).resolve()
            if not folder_path.exists() or not folder_path.is_dir():
                return {"ok": False, "error": f"Pasta inválida: {folder_path}"}

            template_mode = str(payload.get("templateMode") or "PADRAO").upper()
            ds_beline = bool(payload.get("dsBeline", False))

            # opcional (se você quiser permitir o front mandar o caminho final)
            out_path_in = str(payload.get("outPath") or "").strip()
            if out_path_in:
                out_path = Path(out_path_in).resolve()
            else:
                out_path = folder_path / f"preenchido_{_timestamp()}.xlsx"

            # ====== NOVO: Fluxo DS Beline (ROBO + checkbox) ======
            if template_mode == "ROBO" and ds_beline:
                base_xlsx_in = str(payload.get("baseXlsxPath") or payload.get("templatePath") or "").strip()
                if not base_xlsx_in:
                    return {"ok": False, "error": "baseXlsxPath vazio (selecione o Excel do fornecedor)"}

                base_xlsx = Path(base_xlsx_in).resolve()
                if not base_xlsx.exists() or not base_xlsx.is_file():
                    return {"ok": False, "error": f"Excel base inválido: {base_xlsx}"}

                parceiro_template = (self.base_dir / "docs" / "modelo_excel_parceiro.xlsx").resolve()
                if not parceiro_template.exists():
                    return {"ok": False, "error": f"Modelo parceiro não encontrado: {parceiro_template}"}

                ok, problemas = preencher_excel_parceiro(
                    template_path=parceiro_template,
                    pasta_raiz=folder_path,
                    output_path=out_path,
                    base_xlsx_path=base_xlsx,
                )

                return {
                    "ok": True,
                    "outPath": str(out_path),
                    "linhas_ok": ok,
                    "linhas_problemas": problemas,
                }

            # ====== Fluxo normal (o seu atual) ======
            if template_mode == "PADRAO":
                template_path = (self.base_dir / "docs" / "modelo_excel.xlsx").resolve()
            elif template_mode == "ROBO":
                template_path = (self.base_dir / "docs" / "modelo_excel_robo.xlsx").resolve()
            elif template_mode == "NOVO":
                template_path = Path(str(payload.get("templatePath") or "")).resolve()
            else:
                return {"ok": False, "error": f"Modo inválido: {template_mode}"}

            if not template_path.exists():
                return {"ok": False, "error": f"Modelo Excel não encontrado: {template_path}"}

            ok, problemas = preencher_excel(template_path, folder_path, out_path)
            return {"ok": True, "outPath": str(out_path), "linhas_ok": ok, "linhas_problemas": problemas}

        except Exception as e:
            return {"ok": False, "error": str(e)}


    # -------------------------------------------------------------------------
    # AÇÃO: Renomear pastas (CCS / Parceiro / Sequencial) com prévia/execução
    # -------------------------------------------------------------------------
    def run_renomear(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            folder_path = Path(str(payload.get("folderPath") or "")).resolve()
            if not folder_path.exists() or not folder_path.is_dir():
                return {"ok": False, "error": f"Pasta inválida: {folder_path}"}

            dry_run = bool(payload.get("dryRun", True))
            ds_beline = bool(payload.get("dsBeline", False))
            somente_ccs = bool(payload.get("somenteCcs", False))
            limpar_numeracao = bool(payload.get("limparNumeracao", False))

            # ===== LIMPAR NUMERACAO =====
            if limpar_numeracao:
                plan = build_plan_limpar_numeracao(folder_path)
                return _safe_rename_plan(plan, dry_run=dry_run)

            # ===== DS BELINE =====
            if ds_beline:
                xlsx_in = str(payload.get("xlsxPath") or "").strip()
                if not xlsx_in:
                    return {"ok": False, "error": "xlsxPath vazio (selecione o Excel do parceiro)"}
                excel_path = Path(xlsx_in).resolve()
                if not excel_path.exists() or not excel_path.is_file():
                    return {"ok": False, "error": f"Excel inválido: {excel_path}"}

                logs: List[str] = []
                def _log(msg: str) -> None:
                    logs.append(msg)

                plan = build_plan_codigo_beline(folder_path, excel_path, log=_log)
                if not plan:
                    return {"ok": True, "renamed": 0, "skipped": 0, "errors": 0, "logs": logs}

                result = _safe_rename_plan(plan, dry_run=dry_run)

                # junta logs do builder + logs do rename
                result["logs"] = logs + result.get("logs", [])
                return result

            # ===== SOMENTE CCS =====
            if somente_ccs:
                raw_ccs = str(payload.get("ccsStart") or "").strip()
                if not raw_ccs:
                    return {"ok": False, "error": "ccsStart vazio (Somente CCS precisa do número inicial)"}
                ccs_start = int(raw_ccs)

                plan = build_plan_somente_css(folder_path, ccs_start)
                return _safe_rename_plan(plan, dry_run=dry_run)


            # ===== NORMAL (CCS/PARCEIRO) =====
            raw_ccs = str(payload.get("ccsStart") or "").strip()
            raw_parceiro = str(payload.get("parceiroStart") or "").strip()
            if not raw_ccs and not raw_parceiro:
                return {"ok": False, "error": "Preencha Número CCS ou Número Parceiro."}

            ccs_start = int(raw_ccs) if raw_ccs else None
            parceiro_start = int(raw_parceiro) if raw_parceiro else None

            plan = build_plan_ccs_parceiro_sequencial(folder_path, ccs_start, parceiro_start)
            return _safe_rename_plan(plan, dry_run=dry_run)


        except Exception as e:
            return {"ok": False, "error": str(e)}


    # -------------------------------------------------------------------------
    # AÇÃO: Corrigir prefixo antes do primeiro ponto (prévia/execução)
    # -------------------------------------------------------------------------
    def run_corrigir(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Corrige/insere um prefixo antes do primeiro "." do nome da pasta.

        payload esperado:
          {
            folderPath: "...",
            prefix: "010" (exemplo),
            dryRun: true/false,
            padronizarJudicial: false,
            parceiroJudicial: "UNI ASSESSORIA",
            trocarHifenPorPonto: false,
            trocarPontoPorHifen: false
          }
        """
        try:
            folder_path = Path(str(payload.get("folderPath", ""))).resolve()
            if not folder_path.exists() or not folder_path.is_dir():
                return {"ok": False, "error": f"Pasta base inválida: {folder_path}"}

            prefix = str(payload.get("prefix", "") or "")
            dry_run = bool(payload.get("dryRun", True))
            padronizar_judicial = bool(payload.get("padronizarJudicial", False))
            parceiro_judicial = str(payload.get("parceiroJudicial") or "UNI ASSESSORIA").strip()
            trocar_hifen_por_ponto = bool(payload.get("trocarHifenPorPonto", False))
            trocar_ponto_por_hifen = bool(payload.get("trocarPontoPorHifen", False))

            # Evita prefixo bizarro que vira só "."
            if not (padronizar_judicial or trocar_hifen_por_ponto or trocar_ponto_por_hifen) and prefix.strip() == ".":
                return {"ok": False, "error": "Prefixo não pode ser apenas '.'"}

            plan = build_plan_corrigir_opcoes(
                folder_path,
                prefix,
                padronizar_judicial=padronizar_judicial,
                partner_name=parceiro_judicial,
                trocar_hifen_por_ponto=trocar_hifen_por_ponto,
                trocar_ponto_por_hifen=trocar_ponto_por_hifen,
            )

            if not plan:
                return {"ok": True, "logs": [], "renamed": 0, "skipped": 0, "errors": 0}

            return _safe_rename_plan(plan, dry_run=dry_run)

        except ValueError as e:
            return {"ok": False, "error": str(e), "logs": [], "renamed": 0, "skipped": 0, "errors": 1}
        except Exception as e:
            return {"ok": False, "error": str(e), "logs": [], "renamed": 0, "skipped": 0, "errors": 1}

    # -------------------------------------------------------------------------
    # AÇÃO: Gerador de Lotes (integração) — backend Python
    # -------------------------------------------------------------------------
    def lotes_default_model(self) -> Dict[str, Any]:
        """
        Retorna qual modelo o backend vai usar por padrão.

        Retorno:
          - {"ok": True, "path": "...", "name": "..."}
          - {"ok": False, "error": "..."}
        """
        try:
            return {"ok": True, **lotes_modelo_info()}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def lotes_preview_columns(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Lê o Excel do LOTE e devolve a lista de colunas encontradas.

        Útil para você (no frontend) montar selects/combos ou para debug.

        payload:
          {
            "caminho_lote": "...xlsx",
            "coluna_processo_hint": "Número do Processo" (opcional)
          }
        """
        try:
            caminho = str(payload.get("caminho_lote", "") or "")
            hint = payload.get("coluna_processo_hint")
            cols = listar_colunas_lote(caminho, coluna_processo_hint=hint)
            return {"ok": True, "columns": cols}
        except Exception as e:
            return {"ok": False, "error": str(e), "columns": []}

    def run_lotes(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Gera os arquivos de integração a partir do LOTE.

        payload esperado:
          {
            "caminho_lote": "...xlsx",
            "coluna_processo": "Número do Processo",
            "evento_map": {"CALCP": "...", "HC30%": "...", ...},
            "solicitado_por": "45270",
            "modelo_path": null | "...xlsx" (opcional),
            "output_dir": null | "..." (opcional),
            "eventos": null | ["CALCP", ...] (opcional)
          }

        Retorno:
          - {"ok": True, "output_dir": "...", "files": [...], "warnings": [...], "debug": {...}}
          - {"ok": False, "error": "..."}
        """
        try:
            result = gerar_lotes_dict(payload)
            return {"ok": True, **result}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # -------------------------------------------------------------------------
    # AÇÃO: Gerador de Documentos (integração) — backend Python
    # -------------------------------------------------------------------------
    def gerdoc_modelo_info(self) -> Dict[str, Any]:
        try:
            return {"ok": True, **gerdoc_modelo_info_service()}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_pick_planilha(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            path = w.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("Planilhas (*.xlsx;*.xlsm;*.xls;*.csv)",),
            )

            if not path:
                return {"canceled": True}

            if isinstance(path, (list, tuple)):
                p = path[0] if path else ""
            else:
                p = path

            planilha = Path(str(p)).resolve()
            if not planilha.exists() or not planilha.is_file():
                return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {planilha}"}

            if planilha.suffix.lower() not in {".xlsx", ".xlsm", ".xls", ".csv"}:
                return {"canceled": False, "ok": False, "error": "Selecione uma planilha Excel ou CSV."}

            st = planilha.stat()
            return {
                "canceled": False,
                "ok": True,
                "path": str(planilha),
                "name": planilha.name,
                "size": int(st.st_size),
                "lastModified": int(st.st_mtime * 1000),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_pick_saida(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            path = w.create_file_dialog(webview.FOLDER_DIALOG)

            if not path:
                return {"canceled": True}

            if isinstance(path, (list, tuple)):
                p = path[0] if path else ""
            else:
                p = path

            pasta = Path(str(p)).resolve()
            if not pasta.exists() or not pasta.is_dir():
                return {"canceled": False, "ok": False, "error": f"Pasta inválida: {pasta}"}

            return {"canceled": False, "ok": True, "path": str(pasta), "name": pasta.name}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_validar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            info = gerdoc_validar_payload(
                payload.get("planilha"),
                payload.get("pasta_saida"),
                bool(payload.get("gerar_pdf", False)),
            )
            return {"ok": True, "mensagem": "Validação concluída.", "info": info}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_gerar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            result = gerdoc_iniciar_job(
                payload.get("planilha"),
                payload.get("pasta_saida"),
                bool(payload.get("gerar_pdf", False)),
            )
            return {"ok": True, **result}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_status(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            job_id = payload.get("job_id") if isinstance(payload, dict) else payload
            return {"ok": True, "job": gerdoc_obter_status_job(job_id)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_cancelar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            job_id = payload.get("job_id") if isinstance(payload, dict) else payload
            return {"ok": True, "job": gerdoc_cancelar_job(job_id)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gerdoc_abrir_pasta(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            caminho = payload.get("caminho") if isinstance(payload, dict) else payload
            return {"ok": True, **gerdoc_abrir_pasta_service(caminho)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # -------------------------------------------------------------------------
    # Dialog: selecionar PDFs (múltiplos)
    # -------------------------------------------------------------------------
    def divisorpdf_pick_pdf(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            paths = w.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("PDF (*.pdf)",),
            )

            if not paths:
                return {"canceled": True}

            if isinstance(paths, (list, tuple)):
                path = paths[0] if paths else ""
            else:
                path = paths

            pdf_path = Path(str(path)).resolve()
            if not pdf_path.exists() or not pdf_path.is_file():
                return {"canceled": False, "ok": False, "error": f"Arquivo invalido: {pdf_path}"}
            if pdf_path.suffix.lower() != ".pdf":
                return {"canceled": False, "ok": False, "error": "Selecione apenas arquivos PDF (.pdf)."}

            suggested_output = pdf_path.with_name(f"{pdf_path.stem}_dividido")
            return {
                "canceled": False,
                "ok": True,
                "path": str(pdf_path),
                "name": pdf_path.name,
                "stem": pdf_path.stem,
                "suggested_output": str(suggested_output),
            }
        except Exception as e:
            return {"canceled": False, "ok": False, "error": str(e)}

    def divisorpdf_pick_saida(self) -> Dict[str, Any]:
        try:
            w = self._require_window()
            paths = w.create_file_dialog(webview.FOLDER_DIALOG)

            if not paths:
                return {"canceled": True}

            if isinstance(paths, (list, tuple)):
                path = paths[0] if paths else ""
            else:
                path = paths

            folder_path = Path(str(path)).resolve()
            if not folder_path.exists() or not folder_path.is_dir():
                return {"canceled": False, "ok": False, "error": f"Pasta invalida: {folder_path}"}

            return {
                "canceled": False,
                "ok": True,
                "path": str(folder_path),
                "name": folder_path.name,
            }
        except Exception as e:
            return {"canceled": False, "ok": False, "error": str(e)}

    def divisorpdf_processar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            result = divisorpdf_processar_payload(payload if isinstance(payload, dict) else {})
            return {"ok": True, **result}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def divisorpdf_abrir_pasta(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            caminho = payload.get("caminho") if isinstance(payload, dict) else payload
            return {"ok": True, **divisorpdf_abrir_pasta_service(caminho)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def pdfexcel_pick_files(self) -> Dict[str, Any]:
        """
        Abre um diálogo para selecionar múltiplos arquivos PDF.

        Retorno:
          - {"canceled": True} se cancelar
          - {"canceled": False, "files": ["...pdf", ...]} se ok
        """
        w = self._require_window()
        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("PDF (*.pdf)",),
        )

        if not paths:
            return {"canceled": True}

        # pywebview pode retornar lista/tupla ou string (dependendo da versão)
        if isinstance(paths, (str, bytes)):
            paths = [paths]

        files = [str(Path(p).resolve()) for p in paths]
        return {"canceled": False, "files": files}

    # -------------------------------------------------------------------------
    # (Opcional) API: listar PDFs dentro de uma pasta
    # -------------------------------------------------------------------------
    def pdfexcel_list_folder(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Lista PDFs de uma pasta (com ou sem subpastas)."""
        try:
            folder = str(payload.get("folder_path", "") or "")
            recursive = bool(payload.get("recursive", True))
            files = listar_pdfs_em_pasta(folder, recursive=recursive)
            return {"ok": True, "files": files, "count": len(files)}
        except Exception as e:
            return {"ok": False, "error": str(e), "files": [], "count": 0}

    # -------------------------------------------------------------------------
    # AÇÃO: processar PDFs (extrair dados) — backend Python
    # -------------------------------------------------------------------------
    def pdfexcel_process(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Processa PDFs e retorna registros (rows) para o frontend.

        payload:
          {
            "pdf_paths": ["...pdf", ...] (opcional se folder_path existir)
            "folder_path": "..." (opcional)
            "recursive": true|false (opcional, default True)
          }
        """
        try:
            result = pdfexcel_process_payload_dict(payload)
            return {"ok": True, **result}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # -------------------------------------------------------------------------
    # Dialog: escolher onde salvar o Excel (Salvar como)
    # -------------------------------------------------------------------------
    def pdfexcel_pick_save_xlsx(self) -> Dict[str, Any]:
        """Abre diálogo de 'Salvar como' para escolher o caminho do .xlsx."""
        w = self._require_window()
        path = w.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="dados_lote.xlsx",
            file_types=("Excel (*.xlsx)",),
        )

        if not path:
            return {"canceled": True}

        # pywebview pode retornar string ou lista/tupla (dependendo da versão)
        if isinstance(path, (list, tuple)):
            path = path[0] if path else ""

        return {"canceled": False, "path": str(Path(path).resolve())}

    # -------------------------------------------------------------------------
    # AÇÃO: salvar Excel no caminho escolhido
    # -------------------------------------------------------------------------
    def pdfexcel_save_xlsx(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Salva o Excel com base em linhas já extraídas.

        payload:
          {"output_path": "...xlsx", "rows": [...]}
        """
        try:
            out_path = str(payload.get("output_path", "") or "")
            rows = payload.get("rows") or []

            if not out_path:
                return {"ok": False, "error": "output_path vazio"}

            gerar_excel_em_arquivo_lote(rows, out_path)
            return {"ok": True, "path": out_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # -------------------------------------------------------------------------
    # IA_LAUDO — Dialog: selecionar 1 PDF
    # -------------------------------------------------------------------------
    def ialaudos_pick_pdf(self) -> Dict[str, Any]:
        w = self._require_window()

        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("PDF (*.pdf)",),
        )

        if not paths:
            return {"canceled": True}

        # pywebview pode retornar string ou lista/tupla
        if isinstance(paths, (list, tuple)):
            p = paths[0] if paths else ""
        else:
            p = paths

        pdf_path = Path(str(p)).resolve()

        if not pdf_path.exists() or not pdf_path.is_file():
            return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {pdf_path}"}

        return {"canceled": False, "ok": True, "path": str(pdf_path), "name": pdf_path.name}

    # -------------------------------------------------------------------------
    # IA_LAUDO — AÇÃO: analisar PDF (chama services/IA_LAUDO/backend/main.py)
    # -------------------------------------------------------------------------
    def ialaudos_analisar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            pdf_path_in = str(payload.get("pdf_path", "") or "").strip()
            if not pdf_path_in:
                return {"ok": False, "error": "pdf_path vazio"}

            pdf_path = Path(pdf_path_in).resolve()
            if not pdf_path.exists() or not pdf_path.is_file():
                return {"ok": False, "error": f"PDF inválido: {pdf_path}"}

            texto = ialaudos_analisar_pdf(pdf_path)

            # Se seu main.py retornar "Erro: ..." em falhas, a UI mostra bonitinho
            if isinstance(texto, str) and texto.strip().lower().startswith("erro"):
                return {"ok": False, "error": texto}

            return {"ok": True, "text": texto}
        except Exception as e:
            return {"ok": False, "error": str(e)}


    # -------------------------------------------------------------------------
    # IA_PET_INICIAL — Dialog: selecionar até 6 PDFs
    # -------------------------------------------------------------------------
    def iapet_pick_pdfs(self) -> Dict[str, Any]:
        w = self._require_window()

        paths = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("PDF (*.pdf)",),
        )

        if not paths:
            return {"canceled": True}

        if isinstance(paths, (str, Path)):
            paths = [str(paths)]
        else:
            paths = list(paths)

        files = []
        for p in paths[:6]:
            pdf_path = Path(str(p)).resolve()
            if not pdf_path.exists() or not pdf_path.is_file():
                return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {pdf_path}"}
            if pdf_path.suffix.lower() != ".pdf":
                return {"canceled": False, "ok": False, "error": "Selecione apenas PDFs (.pdf)."}

            st = pdf_path.stat()
            files.append({
                "path": str(pdf_path),
                "name": pdf_path.name,
                "size": int(st.st_size),
                "lastModified": int(st.st_mtime * 1000),
            })

        if not files:
            return {"canceled": False, "ok": False, "error": "Nenhum PDF válido selecionado."}

        return {"canceled": False, "ok": True, "files": files}



    # -------------------------------------------------------------------------
    # IA_PET_INICIAL — AÇÃO: analisar PDFs (chama backend)
    # -------------------------------------------------------------------------
    def iapet_analisar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            pdf_paths = payload.get("pdf_paths", [])
            if not isinstance(pdf_paths, list) or not pdf_paths:
                return {"ok": False, "error": "pdf_paths vazio ou inválido."}

            paths = [Path(str(p)).resolve() for p in pdf_paths]
            texto = iapet_analisar_pdfs(paths)

            if isinstance(texto, str) and texto.strip().lower().startswith("erro"):
                return {"ok": False, "error": texto.strip()}

            return {"ok": True, "text": str(texto or "")}

        except Exception as e:
            return {"ok": False, "error": str(e)}


    # -------------------------------------------------------------------------
    # IA_PET_INICIAL — AÇÃO: gerar Word a partir da RESPOSTA (sem chamar IA de novo)
    # -------------------------------------------------------------------------
    
    def iapet_pick_save_docx(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        w = self._require_window()

        suggested = str(payload.get("suggested_name", "") or "").strip()
        if not suggested:
            suggested = "Pet_inicial.docx"
        if not suggested.lower().endswith(".docx"):
            suggested += ".docx"

        path = w.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=suggested,
            file_types=("Word (*.docx)",),
        )

        if not path:
            return {"canceled": True}

        # pywebview pode retornar string ou lista/tupla
        if isinstance(path, (list, tuple)):
            p = path[0] if path else ""
        else:
            p = path

        save_path = Path(str(p)).resolve()
        return {"canceled": False, "ok": True, "path": str(save_path), "name": save_path.name}


    def iapet_gerar_word(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            raw_text = str(payload.get("raw_text", "") or "").strip()
            if not raw_text:
                return {"ok": False, "error": "raw_text vazio."}

            save_path_in = str(payload.get("save_path", "") or "").strip()
            if not save_path_in:
                return {"ok": False, "error": "save_path vazio."}

            save_path = Path(save_path_in).resolve()

            docx_path = iapet_gerar_word_de_resposta(raw_text, saida_path=save_path)
            return {"ok": True, "path": str(docx_path), "name": docx_path.name}

        except Exception as e:
            return {"ok": False, "error": str(e)}


    # -------------------------------------------------------------------------
# RELATORIO_CONFORMIDADE — pick Excel
# -------------------------------------------------------------------------
    def relconf_pick_excel(self) -> Dict[str, Any]:
        w = self._require_window()

        path = w.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx;*.xlsm;*.xls)",),
        )
        if not path:
            return {"canceled": True}

        if isinstance(path, (list, tuple)):
            p = path[0] if path else ""
        else:
            p = path

        excel_path = Path(str(p)).resolve()
        if not excel_path.exists() or not excel_path.is_file():
            return {"canceled": False, "ok": False, "error": f"Arquivo inválido: {excel_path}"}

        st = excel_path.stat()
        return {
            "canceled": False,
            "ok": True,
            "path": str(excel_path),
            "name": excel_path.name,
            "size": int(st.st_size),
            "lastModified": int(st.st_mtime * 1000),
        }


# -------------------------------------------------------------------------
# RELATORIO_CONFORMIDADE — pick output folder
# -------------------------------------------------------------------------
    def relconf_pick_out_dir(self) -> Dict[str, Any]:
        w = self._require_window()

        path = w.create_file_dialog(webview.FOLDER_DIALOG)
        if not path:
            return {"canceled": True}

        if isinstance(path, (list, tuple)):
            p = path[0] if path else ""
        else:
            p = path

        out_dir = Path(str(p)).resolve()
        if not out_dir.exists() or not out_dir.is_dir():
            return {"canceled": False, "ok": False, "error": f"Pasta inválida: {out_dir}"}

        return {"canceled": False, "ok": True, "path": str(out_dir), "name": out_dir.name}


# -------------------------------------------------------------------------
# RELATORIO_CONFORMIDADE — gerar
# -------------------------------------------------------------------------
    def relconf_gerar(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        try:
            excel_path = str(payload.get("excel_path", "") or "").strip()
            out_dir = str(payload.get("out_dir", "") or "").strip()

            res = relconf_gerar_relatorios(excel_path=excel_path, out_dir=out_dir)
            return res if isinstance(res, dict) else {"ok": True, "result": res}
        except Exception as e:
            return {"ok": False, "error": str(e)}
