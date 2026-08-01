import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "dist"
BUILD = ROOT / "build"


def run(cmd):
    print("\n>>", " ".join(cmd))
    subprocess.check_call(cmd, cwd=str(ROOT))


def clean():
    for folder in [OUTPUT, BUILD]:
        if folder.exists():
            try:
                shutil.rmtree(folder)
            except Exception as e:
                # O OneDrive ou o Windows Explorer costumam bloquear a remoção da pasta raiz.
                # Nesses casos, tentamos remover o conteúdo interno e ignorar a falha na pasta raiz.
                print(f"Aviso: Não foi possível remover a pasta raiz {folder.name} ({e}).")
                print("Limpando arquivos internos...")
                for item in folder.iterdir():
                    try:
                        if item.is_dir():
                            shutil.rmtree(item, ignore_errors=True)
                        else:
                            item.unlink()
                    except Exception as err:
                        print(f"Não foi possível remover {item.name}: {err}")


def main():
    clean()
    run([sys.executable, "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel", "pyinstaller"])
    
    # Lista de argumentos para o PyInstaller para melhor legibilidade
    pyinstaller_args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "splitvision.py",
        # --- Configurações Básicas ---
        "--name", "splitvision",
        "--onefile",
        "--noconsole",
        "--icon",
        "splitvision.ico",
        "--add-data",
        "splitvision.ico;.",
        # --- Caminhos de Saída ---
        "--distpath",
        str(OUTPUT),
        "--workpath",
        str(BUILD),
        # --- Otimizações e Exclusões ---
        "--optimize",
        "2",
        "--exclude-module",
        "matplotlib",
        "--exclude-module",
        "tkinter.test",
        # --- Coleta de Dados e Imports Ocultos ---
        # Coleta automaticamente todos os dados necessários do pypdfium2 e pdf2docx
        "--collect-data", "pypdfium2",
        "--collect-data", "pdf2docx",
        # Imports implícitos (hidden imports) para winrt e outras dependências
        "--hidden-import", "winrt",
        "--hidden-import", "winrt.windows.media.ocr",
        "--hidden-import", "winrt.windows.globalization",
        "--hidden-import", "winrt.windows.storage.streams",
        "--hidden-import", "winrt.windows.graphics.imaging",
        "--hidden-import", "winrt.windows.foundation",
        "--hidden-import", "winrt.windows.foundation.collections",
        "--hidden-import", "winocr",
        "--hidden-import", "windnd",
        "--hidden-import", "pdf2docx",
        "--hidden-import", "docx",
        "--hidden-import", "docx2pdf",
        "--hidden-import", "win32com",
        "--hidden-import", "conversor_pdf_word",
    ]

    run(pyinstaller_args)
    print("\nBuild concluído em:", OUTPUT)


if __name__ == "__main__":
    main()
