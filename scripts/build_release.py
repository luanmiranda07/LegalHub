from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def zip_folder(source: Path, destination: Path) -> None:
    if destination.exists():
        destination.unlink()

    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for path in source.rglob("*"):
            if path.is_file():
                zipf.write(path, path.relative_to(source.parent))


def run(command: list[str]) -> None:
    print("$", " ".join(command))
    subprocess.run(command, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Gera pacote de release do LegalHub.")
    parser.add_argument("--version", required=True, help="Versão semântica. Ex.: 1.0.1")
    parser.add_argument(
        "--base-url",
        required=True,
        help="URL pública da VM sem barra final. Ex.: https://legalhub.seudominio.com",
    )
    parser.add_argument("--name", default="LegalHub", help="Nome do executável principal")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    dist_dir = root / "dist"
    exe_dir = dist_dir / args.name
    releases_dir = root / "releases"
    updates_dir = root / "updates"

    releases_dir.mkdir(exist_ok=True)
    updates_dir.mkdir(exist_ok=True)

    run([
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        args.name,
        "--add-data",
        "src/docs;docs",
        "--add-data",
        "src/web;web",
        "--add-data",
        "src/services;services",
        "src/app.py",
    ])

    if not exe_dir.exists():
        raise FileNotFoundError(f"Pasta de build não encontrada: {exe_dir}")

    zip_path = releases_dir / f"{args.name}-{args.version}.zip"
    zip_folder(exe_dir, zip_path)

    digest = sha256_file(zip_path)
    latest = {
        "version": args.version,
        "url": f"{args.base_url.rstrip('/')}/releases/{zip_path.name}",
        "sha256": digest,
        "mandatory": True,
        "notes": "Nova versão publicada.",
    }

    latest_path = updates_dir / "latest.json"
    latest_path.write_text(json.dumps(latest, ensure_ascii=False, indent=4), encoding="utf-8")

    print("\nRelease gerada com sucesso:")
    print(f"ZIP: {zip_path}")
    print(f"SHA256: {digest}")
    print(f"Manifesto: {latest_path}")
    print("\nEnvie para a VM:")
    print(f"- {zip_path} -> /opt/legalhub/releases/{zip_path.name}")
    print(f"- {latest_path} -> /opt/legalhub/updates/latest.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
