from __future__ import annotations

import re
import threading
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


PROCESSO_REGEXES = [
    re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}"),
    re.compile(r"(?:proc\.?|processo)\s*:?\s*([0-9.\-\/]+)", re.IGNORECASE),
    re.compile(r"([0-9][0-9.\-/]{10,})"),
]

COLUNAS_PROCESSO_CANDIDATAS = {
    "numero do processo principal",
    "número do processo principal",
    "numero_do_processo_principal",
    "número_do_processo_principal",
    "numero do processo",
    "número do processo",
    "numero_do_processo",
    "número_do_processo",
    "processo principal",
    "processo",
}

COLUNAS_CHANCE_CANDIDATAS = {
    "chance_de_exito",
    "chance de exito",
    "chance_de_exito_",
    "chance de êxito",
    "chance_de_êxito",
}

EXTENSOES_ARQUIVO = {".pdf"}


@dataclass(slots=True)
class MatchItem:
    caminho: Path
    processo_extraido: str | None
    processo_normalizado: str | None
    chance_exito: str | None
    novo_nome: str | None
    status: str


def remover_acentos(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch))


def normalizar_texto(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = remover_acentos(texto).strip().lower()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_chave_coluna(valor: object) -> str:
    texto = normalizar_texto(valor)
    texto = texto.replace("ç", "c")
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def limpar_nome_arquivo(valor: str, limite: int = 100) -> str:
    texto = remover_acentos(valor)
    texto = re.sub(r'[\\/:*?"<>|]', " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip(" .-")
    return texto[:limite].rstrip(" .-") or "SEM_VALOR"


def detectar_coluna(df: pd.DataFrame, candidatas: set[str], contem: str | None = None) -> str | None:
    candidatas_normalizadas = {normalizar_chave_coluna(col) for col in candidatas}
    mapa = {normalizar_chave_coluna(col): col for col in df.columns}

    for chave_normalizada, coluna_original in mapa.items():
        if chave_normalizada in candidatas_normalizadas:
            return coluna_original

    if contem:
        contem_normalizado = normalizar_chave_coluna(contem)
        for chave_normalizada, coluna_original in mapa.items():
            if contem_normalizado in chave_normalizada:
                return coluna_original

    return None


def normalizar_numero_processo(valor: object) -> str:
    if valor is None:
        return ""

    texto = str(valor).strip()

    if not texto:
        return ""

    texto = texto.replace("\u00A0", " ")
    texto = re.sub(r"\s+", "", texto)

    # Se vier como número com .0 no Excel, remove
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".", 1)[0]

    return re.sub(r"\D+", "", texto)


def carregar_mapa_excel(caminho_excel: Path) -> tuple[dict[str, str], str, str]:
    try:
        df = pd.read_excel(caminho_excel)
    except ImportError as exc:
        raise ValueError(
            "Nao foi possivel ler este Excel com as dependencias instaladas. "
            "Salve a planilha como .xlsx e tente novamente."
        ) from exc

    if df.empty:
        raise ValueError("O Excel está vazio.")

    coluna_processo = detectar_coluna(df, COLUNAS_PROCESSO_CANDIDATAS, contem="processo")
    coluna_chance = detectar_coluna(df, COLUNAS_CHANCE_CANDIDATAS, contem="chance")

    if not coluna_processo:
        raise ValueError(
            "Não encontrei a coluna do processo no Excel. "
            "Use um nome como: NÚMERO DO PROCESSO PRINCIPAL ou NÚMERO DO PROCESSO."
        )

    if not coluna_chance:
        raise ValueError(
            "Não encontrei a coluna 'CHANCE DE ÊXITO' no Excel. "
            "Use um nome como: CHANCE DE ÊXITO ou chance_de_exito."
        )

    mapa: dict[str, str] = {}

    for _, linha in df[[coluna_processo, coluna_chance]].iterrows():
        processo = normalizar_numero_processo(linha[coluna_processo])
        chance = "" if pd.isna(linha[coluna_chance]) else str(linha[coluna_chance]).strip()

        if not processo or not chance:
            continue

        if processo not in mapa:
            mapa[processo] = chance

    if not mapa:
        raise ValueError("Nenhum número de processo com chance de êxito válida foi encontrado no Excel.")

    return mapa, coluna_processo, coluna_chance


def extrair_numero_processo(nome_base: str) -> str | None:
    for regex in PROCESSO_REGEXES:
        match = regex.search(nome_base)
        if not match:
            continue

        valor = match.group(1) if match.lastindex else match.group(0)
        processo = normalizar_numero_processo(valor)

        if len(processo) >= 10:
            return processo

    return None


def extrair_faixa_processo(nome_base: str) -> tuple[int, int] | None:
    for regex in PROCESSO_REGEXES:
        match = regex.search(nome_base)
        if match:
            if match.lastindex:
                return match.span(1)
            return match.span(0)
    return None


def montar_novo_nome(nome_atual: str, faixa_processo: tuple[int, int] | None, chance_exito: str) -> str:
    chance_limpa = limpar_nome_arquivo(chance_exito)
    stem = Path(nome_atual).stem
    suffix = Path(nome_atual).suffix

    if faixa_processo:
        inicio, fim = faixa_processo
        novo_stem = f"{stem[:fim]} - {chance_limpa}{stem[fim:]}"
    else:
        novo_stem = f"{stem} - {chance_limpa}"

    novo_stem = re.sub(r"\s+", " ", novo_stem).strip()
    novo_stem = re.sub(r"\s+-\s+-\s+", " - ", novo_stem)
    return f"{novo_stem}{suffix}"


def nome_ja_contem_chance(nome_atual: str, chance_exito: str) -> bool:
    stem_normalizado = normalizar_texto(Path(nome_atual).stem)
    chance_normalizada = normalizar_texto(limpar_nome_arquivo(chance_exito))
    return bool(chance_normalizada and chance_normalizada in stem_normalizado)


def garantir_nome_unico(destino: Path) -> Path:
    if not destino.exists():
        return destino

    contador = 1
    while True:
        candidato = destino.with_name(f"{destino.stem} ({contador}){destino.suffix}")
        if not candidato.exists():
            return candidato
        contador += 1


def iterar_itens(pasta: Path) -> list[Path]:
    return sorted(
        [
            caminho
            for caminho in pasta.rglob("*")
            if caminho.is_file() and caminho.suffix.lower() in EXTENSOES_ARQUIVO
        ],
        key=lambda item: str(item).lower(),
    )


def analisar_itens(pasta: Path, mapa_excel: dict[str, str]) -> list[MatchItem]:
    resultados: list[MatchItem] = []

    for caminho in iterar_itens(pasta):
        nome_base = caminho.stem if caminho.is_file() else caminho.name
        processo_extraido = extrair_numero_processo(nome_base)
        processo_normalizado = processo_extraido if processo_extraido else None
        faixa_processo = extrair_faixa_processo(nome_base)

        if not processo_extraido or not processo_normalizado:
            resultados.append(
                MatchItem(
                    caminho=caminho,
                    processo_extraido=None,
                    processo_normalizado=None,
                    chance_exito=None,
                    novo_nome=None,
                    status="Número do processo não encontrado no nome do item.",
                )
            )
            continue

        chance_exito = mapa_excel.get(processo_normalizado)
        if not chance_exito:
            resultados.append(
                MatchItem(
                    caminho=caminho,
                    processo_extraido=processo_extraido,
                    processo_normalizado=processo_normalizado,
                    chance_exito=None,
                    novo_nome=None,
                    status="Número do processo não encontrado no Excel.",
                )
            )
            continue

        if nome_ja_contem_chance(caminho.name, chance_exito):
            resultados.append(
                MatchItem(
                    caminho=caminho,
                    processo_extraido=processo_extraido,
                    processo_normalizado=processo_normalizado,
                    chance_exito=chance_exito,
                    novo_nome=caminho.name,
                    status="Nome jÃ¡ estÃ¡ atualizado.",
                )
            )
            continue

        novo_nome = montar_novo_nome(caminho.name, faixa_processo, chance_exito)
        if novo_nome == caminho.name:
            resultados.append(
                MatchItem(
                    caminho=caminho,
                    processo_extraido=processo_extraido,
                    processo_normalizado=processo_normalizado,
                    chance_exito=chance_exito,
                    novo_nome=novo_nome,
                    status="Nome já está atualizado.",
                )
            )
            continue

        resultados.append(
            MatchItem(
                caminho=caminho,
                processo_extraido=processo_extraido,
                processo_normalizado=processo_normalizado,
                chance_exito=chance_exito,
                novo_nome=novo_nome,
                status="OK",
            )
        )

    return resultados


def aplicar_renomeacao(resultados: list[MatchItem]) -> tuple[int, int]:
    sucesso = 0
    erro = 0

    for item in resultados:
        if item.status != "OK" or not item.novo_nome:
            continue

        origem = item.caminho
        destino = garantir_nome_unico(origem.with_name(item.novo_nome))

        try:
            origem.rename(destino)
            sucesso += 1
        except Exception:
            erro += 1

    return sucesso, erro


def validar_payload(payload: dict[str, Any]) -> tuple[Path, Path]:
    pasta_raw = str(payload.get("pasta") or "").strip()
    excel_raw = str(payload.get("excel") or "").strip()

    if not pasta_raw:
        raise ValueError("Selecione uma pasta antes de continuar.")

    if not excel_raw:
        raise ValueError("Selecione um arquivo Excel antes de continuar.")

    pasta = Path(pasta_raw).expanduser().resolve()
    excel = Path(excel_raw).expanduser().resolve()

    if not pasta.exists() or not pasta.is_dir():
        raise ValueError(f"Pasta invalida: {pasta}")

    if not excel.exists() or not excel.is_file():
        raise ValueError(f"Arquivo Excel invalido: {excel}")

    if excel.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("Selecione uma planilha .xlsx, .xlsm ou .xls.")

    return pasta, excel


def _logs_resultados(
    resultados: list[MatchItem],
    coluna_processo: str,
    coluna_chance: str,
    *,
    dry_run: bool,
    sucesso: int = 0,
    erro_renomeacao: int = 0,
) -> list[str]:
    logs = [
        f"Coluna de processo usada: {coluna_processo}",
        f"Coluna de chance usada: {coluna_chance}",
        "-" * 90,
    ]

    if not resultados:
        logs.append("Nenhum PDF encontrado na pasta selecionada.")

    for item in resultados:
        if item.status == "OK":
            prefixo = "[PREVIA]" if dry_run else "[RENOMEADO]"
            logs.append(f"{prefixo} {item.caminho.name}")
            logs.append(f"          Processo: {item.processo_extraido}")
            logs.append(f"          Chance: {item.chance_exito}")
            logs.append(f"          -> {item.novo_nome}")
        else:
            prefixo = "[AVISO]" if dry_run else "[IGNORADO]"
            logs.append(f"{prefixo} {item.caminho.name} | {item.status}")

    prontos = sum(1 for item in resultados if item.status == "OK")
    ignorados = len(resultados) - prontos

    logs.append("-" * 90)
    if dry_run:
        logs.append(f"Prontos para renomear: {prontos} | Ignorados/avisos: {ignorados}")
    else:
        logs.append(f"Renomeados: {sucesso} | Ignorados: {ignorados} | Erros ao renomear: {erro_renomeacao}")

    return logs


def processar_payload(payload: dict[str, Any], *, dry_run: bool) -> dict[str, Any]:
    try:
        pasta, excel = validar_payload(payload)
        mapa_excel, coluna_processo, coluna_chance = carregar_mapa_excel(excel)
        resultados = analisar_itens(pasta, mapa_excel)
        prontos = sum(1 for item in resultados if item.status == "OK")
        ignorados = len(resultados) - prontos
        renomeados = 0
        erros = 0

        if not dry_run:
            renomeados, erros = aplicar_renomeacao(resultados)

        logs = _logs_resultados(
            resultados,
            coluna_processo,
            coluna_chance,
            dry_run=dry_run,
            sucesso=renomeados,
            erro_renomeacao=erros,
        )

        return {
            "ok": erros == 0,
            "mensagem": "Validacao concluida." if dry_run else "Processamento concluido.",
            "logs": logs,
            "total": len(resultados),
            "prontos": prontos,
            "ignorados": ignorados,
            "renomeados": renomeados,
            "erros": erros,
        }
    except Exception as exc:
        return {
            "ok": False,
            "mensagem": str(exc),
            "error": str(exc),
            "logs": [f"[ERRO] {exc}"],
            "total": 0,
            "prontos": 0,
            "ignorados": 0,
            "renomeados": 0,
            "erros": 1,
        }


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Adicionar Chance de Êxito")
        self.root.geometry("860x520")
        self.root.resizable(False, False)

        self.pasta_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Selecione a pasta e o Excel.")

        self._montar_interface()

    def _montar_interface(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Pasta:").place(x=0, y=6)
        ttk.Entry(frame, textvariable=self.pasta_var).place(x=48, y=4, width=690)
        ttk.Button(frame, text="Selecionar", command=self.selecionar_pasta).place(x=750, y=2, width=90)

        ttk.Label(frame, text="Excel:").place(x=0, y=46)
        ttk.Entry(frame, textvariable=self.excel_var).place(x=48, y=44, width=690)
        ttk.Button(frame, text="Selecionar", command=self.selecionar_excel).place(x=750, y=42, width=90)

        ttk.Button(frame, text="Validar", command=self.validar).place(x=250, y=96, width=150)
        ttk.Button(frame, text="Executar", command=self.executar).place(x=430, y=96, width=150)

        ttk.Label(frame, textvariable=self.status_var).place(x=0, y=142)

        self.log = tk.Text(frame, height=17, width=100, wrap="word")
        self.log.place(x=0, y=170, width=822, height=300)
        self.log.configure(state="disabled")

    def selecionar_pasta(self) -> None:
        pasta = filedialog.askdirectory(title="Selecione a pasta")
        if pasta:
            self.pasta_var.set(pasta)

    def selecionar_excel(self) -> None:
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls *.xlsb")],
        )
        if arquivo:
            self.excel_var.set(arquivo)

    def escrever_log(self, texto: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def limpar_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _validar_caminhos(self) -> tuple[Path, Path] | None:
        pasta = Path(self.pasta_var.get().strip())
        excel = Path(self.excel_var.get().strip())

        if not pasta.exists() or not pasta.is_dir():
            messagebox.showerror("Erro", "Selecione uma pasta válida.")
            return None

        if not excel.exists() or not excel.is_file():
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return None

        return pasta, excel

    def validar(self) -> None:
        dados = self._validar_caminhos()
        if not dados:
            return

        pasta, excel = dados
        self.status_var.set("Validando...")
        self.limpar_log()

        def tarefa() -> None:
            try:
                mapa_excel, coluna_processo, coluna_chance = carregar_mapa_excel(excel)
                resultados = analisar_itens(pasta, mapa_excel)
                ok = sum(1 for r in resultados if r.status == "OK")
                observacao = len(resultados) - ok

                self.root.after(
                    0,
                    lambda: self._mostrar_validacao(resultados, coluna_processo, coluna_chance, ok, observacao),
                )
            except Exception as exc:
                self.root.after(0, lambda: self._mostrar_erro(exc))

        threading.Thread(target=tarefa, daemon=True).start()

    def _mostrar_validacao(
        self,
        resultados: list[MatchItem],
        coluna_processo: str,
        coluna_chance: str,
        ok: int,
        observacao: int,
    ) -> None:
        self.status_var.set("Validação concluída.")
        self.escrever_log(f"Coluna de processo usada: {coluna_processo}")
        self.escrever_log(f"Coluna de chance usada: {coluna_chance}")
        self.escrever_log("-" * 90)

        for item in resultados:
            tipo = "PASTA" if item.caminho.is_dir() else "ARQUIVO"
            if item.status == "OK":
                self.escrever_log(f"[OK][{tipo}] {item.caminho.name}")
                self.escrever_log(f"            Processo: {item.processo_extraido}")
                self.escrever_log(f"            Chance: {item.chance_exito}")
                self.escrever_log(f"            -> {item.novo_nome}")
            else:
                self.escrever_log(f"[AVISO][{tipo}] {item.caminho.name} | {item.status}")

        self.escrever_log("-" * 90)
        self.escrever_log(f"Prontos para renomear: {ok} | Com observação: {observacao}")

    def executar(self) -> None:
        dados = self._validar_caminhos()
        if not dados:
            return

        pasta, excel = dados
        self.status_var.set("Executando...")
        self.limpar_log()

        def tarefa() -> None:
            try:
                mapa_excel, coluna_processo, coluna_chance = carregar_mapa_excel(excel)
                resultados = analisar_itens(pasta, mapa_excel)
                sucesso, erro_renomeacao = aplicar_renomeacao(resultados)

                self.root.after(
                    0,
                    lambda: self._mostrar_execucao(
                        resultados,
                        coluna_processo,
                        coluna_chance,
                        sucesso,
                        erro_renomeacao,
                    ),
                )
            except Exception as exc:
                self.root.after(0, lambda: self._mostrar_erro(exc))

        threading.Thread(target=tarefa, daemon=True).start()

    def _mostrar_execucao(
        self,
        resultados: list[MatchItem],
        coluna_processo: str,
        coluna_chance: str,
        sucesso: int,
        erro_renomeacao: int,
    ) -> None:
        self.status_var.set("Processamento concluído.")
        self.escrever_log(f"Coluna de processo usada: {coluna_processo}")
        self.escrever_log(f"Coluna de chance usada: {coluna_chance}")
        self.escrever_log("-" * 90)

        for item in resultados:
            tipo = "PASTA" if item.caminho.is_dir() else "ARQUIVO"
            if item.status == "OK":
                self.escrever_log(f"[RENOMEADO][{tipo}] {item.caminho.name}")
                self.escrever_log(f"                  -> {item.novo_nome}")
            else:
                self.escrever_log(f"[IGNORADO][{tipo}] {item.caminho.name} | {item.status}")

        ignorados = sum(1 for r in resultados if r.status != "OK")
        self.escrever_log("-" * 90)
        self.escrever_log(
            f"Renomeados: {sucesso} | Ignorados: {ignorados} | Erros ao renomear: {erro_renomeacao}"
        )
        messagebox.showinfo(
            "Concluído",
            f"Renomeados: {sucesso}\nIgnorados: {ignorados}\nErros ao renomear: {erro_renomeacao}",
        )

    def _mostrar_erro(self, exc: Exception) -> None:
        self.status_var.set("Erro durante o processamento.")
        self.escrever_log(f"[ERRO] {exc}")
        messagebox.showerror("Erro", str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    estilo = ttk.Style()
    try:
        estilo.theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()
