# -*- mode: python ; coding: utf-8 -*-
"""
Receita do PyInstaller para gerar o executável "TicketsTaxis.exe".

Gera um único arquivo .exe (sem console) com tudo embutido:
  - Tesseract-OCR (tesseract.exe + DLLs + tessdata com por/eng/osd)
  - Python + OpenCV + numpy + Pillow + pytesseract + openpyxl + Tkinter

Gerar com:  pyinstaller --noconfirm TicketsTaxis.spec
(ou execute build_exe.ps1)
"""

import os
from pathlib import Path

RAIZ = Path(SPECPATH).resolve()  # pasta do .spec


# --------------------------------------------------------------------------- #
# Localiza a instalação do Tesseract no computador de build
# --------------------------------------------------------------------------- #

def _achar_tesseract() -> Path | None:
    candidatos = [
        Path(r"C:\Program Files\Tesseract-OCR"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR"),
        RAIZ / "Tesseract-OCR",
    ]
    env = os.environ.get("TESSERACT_EXE")
    if env and Path(env).is_file():
        candidatos.insert(0, Path(env).resolve().parent)

    for pasta in candidatos:
        if (pasta / "tesseract.exe").is_file():
            return pasta
    return None


pasta_tesseract = _achar_tesseract()
datas: list[tuple[str, str]] = []

if pasta_tesseract:
    print(f"[spec] Tesseract embutido de: {pasta_tesseract}")

    # tesseract.exe + todas as DLLs (exe/dll da raiz e de subpastas)
    for item in pasta_tesseract.rglob("*"):
        if not item.is_file():
            continue
        if item.suffix.lower() not in (".dll", ".exe"):
            continue
        rel = item.relative_to(pasta_tesseract)
        sub = "" if str(rel.parent) == "." else "/" + str(rel.parent).replace("\\", "/")
        datas.append((str(item), "Tesseract-OCR" + sub))

    # Somente os idiomas usados (mantém o .exe enxuto)
    for idioma in ("por", "eng", "osd"):
        arq = pasta_tesseract / "tessdata" / f"{idioma}.traineddata"
        if arq.is_file():
            datas.append((str(arq), "Tesseract-OCR/tessdata"))
else:
    print("[spec] AVISO: Tesseract nao encontrado. O .exe dependera de uma")
    print("[spec] instalacao do Tesseract no computador onde for executado.")

# Remove duplicatas preservando a ordem
_vistos = set()
_unicas = []
for origem, destino in datas:
    chave = (os.path.normcase(os.path.abspath(origem)), destino)
    if chave in _vistos:
        continue
    _vistos.add(chave)
    _unicas.append((origem, destino))
datas = _unicas

ocultos = [
    "pytesseract",
    "openpyxl",
    "openpyxl.cell._writer",
    "cv2",
    "numpy",
    "PIL.Image",
    "PIL.ImageOps",
    "PIL._tkinter_finder",
]

a = Analysis(
    ["app_tickets.py"],
    pathex=[str(RAIZ)],
    binaries=[],
    datas=datas,
    hiddenimports=ocultos,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "pandas", "scipy", "IPython", "notebook",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",
        "test", "unittest", "pydoc_data",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TicketsTaxis",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # <- sem janela de terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
