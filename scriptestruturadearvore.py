#!/usr/bin/env python3
import argparse
from pathlib import Path

# Pastas/arquivos ignorados por padrão
DEFAULT_IGNORE = {
    ".git",
    ".idea",
    ".vscode",
    "__pycache__",
    ".mypy_cache",
    ".pytest_cache",
    "venv",
    ".venv",
    "env",
    "node_modules",
    "dist",
    "build",
    "scriptestruturadearvore.py"
}


def build_tree(root: Path, prefix: str, ignore: set[str]) -> list[str]:
    """
    Monta a árvore de diretórios/arquivos recursivamente,
    retornando uma lista de linhas de texto.
    """
    try:
        entries = [e for e in root.iterdir() if e.name not in ignore]
    except PermissionError:
        # Caso não tenha permissão para alguma pasta
        return [f"{prefix}└── [PERMISSION DENIED]"]

    # Ordena: pastas primeiro, depois arquivos, tudo em ordem alfabética
    entries.sort(key=lambda p: (p.is_file(), p.name.lower()))

    lines = []
    total = len(entries)

    for index, path in enumerate(entries):
        is_last = index == (total - 1)
        connector = "└── " if is_last else "├── "

        # Linha atual
        lines.append(f"{prefix}{connector}{path.name}")

        # Se for diretório, desce recursivamente
        if path.is_dir():
            extension = "    " if is_last else "│   "
            lines.extend(build_tree(path, prefix + extension, ignore))

    return lines


def tree_to_md(
    root_dir: str, output_file: str, titulo: str | None, ignore_extra: list[str]
):
    root_path = Path(root_dir).resolve()
    if not root_path.exists():
        raise FileNotFoundError(f"Diretório não encontrado: {root_path}")

    titulo = titulo or root_path.name

    # Monta o set de ignore: default + extras passados por CLI
    ignore = set(DEFAULT_IGNORE)
    ignore.update(name.strip() for name in ignore_extra if name.strip())

    # Cabeçalho em Markdown
    linhas = [
        f"# Estrutura do projeto `{titulo}`",
        "",
        "```text",
        root_path.name,
    ]

    # Corpo da tree
    linhas.extend(build_tree(root_path, "", ignore))

    # Fecha o bloco de código
    linhas.append("```")

    # Salva no arquivo .md
    output_path = Path(output_file)
    output_path.write_text("\n".join(linhas), encoding="utf-8")

    print(f"Árvore gerada em: {output_path}")
    if ignore:
        print("Ignorando:", ", ".join(sorted(ignore)))


def main():
    parser = argparse.ArgumentParser(
        description="Gera uma árvore de diretórios/arquivos em formato Markdown."
    )
    parser.add_argument(
        "root",
        nargs="?",
        default=".",
        help="Diretório raiz do projeto (padrão: diretório atual).",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="TREE.md",
        help="Nome do arquivo de saída .md (padrão: TREE.md).",
    )
    parser.add_argument(
        "-t",
        "--title",
        default=None,
        help="Título do projeto para aparecer no .md (padrão: nome da pasta raiz).",
    )
    parser.add_argument(
        "-i",
        "--ignore",
        default="",
        help=(
            "Lista extra de nomes de pastas/arquivos para ignorar, "
            "separados por vírgula. Ex: '.git,venv,node_modules'"
        ),
    )

    args = parser.parse_args()
    ignore_extra = [x for x in args.ignore.split(",")] if args.ignore else []
    tree_to_md(args.root, args.output, args.title, ignore_extra)


if __name__ == "__main__":
    main()
