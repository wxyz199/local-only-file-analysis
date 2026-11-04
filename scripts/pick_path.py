#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pick local file paths with a native dialog and print them, one per line.

Nothing is uploaded: the dialog only reads paths from the local filesystem.
Two backends are tried in order — tkinter (cross-platform), then the Windows
PowerShell OpenFileDialog (for Python builds shipped without tkinter).

Usage:
  python pick_path.py [--start DIR] [--single] [--backend auto|tk|ps]

Exit codes:
  0  paths printed
  1  user cancelled
  2  no dialog backend available — ask the user to paste an absolute path

The dialog blocks until the user picks files or cancels; do not run it in a
non-interactive context.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile


def _tk_pick(start: str, single: bool) -> list[str] | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if single:
            chosen = filedialog.askopenfilename(title="选择要在本地分析的文件", initialdir=start)
            return [chosen] if chosen else []
        return list(filedialog.askopenfilenames(title="选择要在本地分析的文件（可多选）",
                                                initialdir=start) or [])
    except Exception:  # noqa: BLE001 - TclError when no GUI session
        return None


def _ps_pick(start: str, single: bool) -> list[str] | None:
    if os.name != "nt":
        return None
    tmp = os.path.join(tempfile.gettempdir(), "lof_pick_%d.txt" % os.getpid())
    multiselect = "$false" if single else "$true"
    script = (
        "Add-Type -AssemblyName System.Windows.Forms | Out-Null;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        "$d.Title = 'Select files to analyse locally';"
        f"$d.Multiselect = {multiselect};"
        "$d.Filter = 'Common|*.xlsx;*.xlsm;*.csv;*.tsv;*.docx;*.docm;*.pptx;*.pdf;"
        "*.txt;*.md;*.json;*.html;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|All files|*.*';"
        f"$d.InitialDirectory = '{start.replace(chr(39), chr(39) * 2)}';"
        "$r = $d.ShowDialog();"
        "if ($r -eq [System.Windows.Forms.DialogResult]::OK) {"
        "$lines = @('__OK__') + $d.FileNames } else { $lines = @('__CANCEL__') };"
        f"[System.IO.File]::WriteAllLines('{tmp}', $lines, "
        "[System.Text.Encoding]::UTF8)"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", script],
                       capture_output=True, timeout=600, check=False)
    except (OSError, subprocess.SubprocessError):
        return None
    if not os.path.isfile(tmp):
        return None
    try:
        with open(tmp, encoding="utf-8-sig") as fh:
            lines = [line.strip() for line in fh if line.strip()]
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    if not lines or lines[0] != "__OK__":
        return []
    return lines[1:]


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    start, single, backend = os.path.expanduser("~"), False, "auto"
    args = argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--start" and i + 1 < len(args):
            start = args[i + 1]; i += 2
        elif a == "--single":
            single = True; i += 1
        elif a == "--backend" and i + 1 < len(args):
            backend = args[i + 1]; i += 2
        elif a in ("-h", "--help"):
            print(__doc__); return 0
        else:
            i += 1

    order = {"auto": ("tk", "ps"), "tk": ("tk",), "ps": ("ps",)}.get(backend, ("tk", "ps"))
    for name in order:
        paths = _tk_pick(start, single) if name == "tk" else _ps_pick(start, single)
        if paths is None:
            continue
        if not paths:
            print("未选择文件。若没有弹出选择框，请直接提供文件的本地绝对路径。")
            return 1
        for p in paths:
            print(os.path.abspath(p))
        return 0

    print("ERROR: 本机没有可用的文件选择框（缺少 tkinter，且 PowerShell 弹窗不可用）。"
          "请直接提供文件的本地绝对路径，例如 D:\\资料\\台账.xlsx。", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
