"""Extra opcional (solo macOS con Microsoft Word): abre el .docx, actualiza el índice y guarda."""
from __future__ import annotations

import platform
import subprocess
from pathlib import Path

_SCRIPT = """
on run argv
    set inPath to item 1 of argv
    tell application "Microsoft Word"
        open POSIX file inPath
        set d to active document
        try
            update table of contents 1 of d
        end try
        repaginate d
        try
            update table of contents 1 of d
        end try
        save d
        close d saving yes
    end tell
end run
"""


def actualizar_con_word(docx_path: str | Path) -> None:
    if platform.system() != "Darwin":
        raise RuntimeError("--actualizar-indice solo está disponible en macOS con Word instalado. "
                           "En Windows, abre el .docx y acepta «actualizar campos» al abrir.")
    r = subprocess.run(["osascript", "-e", _SCRIPT, str(Path(docx_path).resolve())],
                       capture_output=True, text=True, timeout=300)
    if r.returncode:
        raise RuntimeError(f"Word devolvió un error: {r.stderr.strip()}")
