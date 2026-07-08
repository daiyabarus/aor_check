"""
Build script -- packages audit_summary_tool.py into a single-file
Windows/Mac/Linux executable using PyInstaller.

Usage:
    pip install -r requirements.txt
    python build.py

Icon:
    If a file named "icon.ico" exists in this same folder, it is used as
    the executable's icon. Otherwise PyInstaller's own default application
    icon is used (no --icon flag is passed).

Output:
    dist/AuditSummaryTool(.exe)
"""

import os
import sys
import shutil

import PyInstaller.__main__

APP_NAME = "AuditSummaryTool"
SCRIPT = "audit_summary_tool.py"
ICON_FILE = "icon.ico"

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    script_path = os.path.join(HERE, SCRIPT)
    if not os.path.isfile(script_path):
        sys.exit(f"Cannot find {SCRIPT} next to build.py ({script_path})")

    args = [
        script_path,
        "--name", APP_NAME,
        "--onefile",
        "--windowed",      # no console window (this is a GUI app)
        "--noconfirm",
        "--clean",
        "--distpath", os.path.join(HERE, "dist"),
        "--workpath", os.path.join(HERE, "build"),
        "--specpath", HERE,
    ]

    icon_path = os.path.join(HERE, ICON_FILE)
    if os.path.isfile(icon_path):
        args += ["--icon", icon_path]
        print(f"Using custom icon: {ICON_FILE}")
    else:
        print(f"No '{ICON_FILE}' found next to build.py -- "
              f"using PyInstaller's default application icon.")

    print("Running PyInstaller with args:", args)
    PyInstaller.__main__.run(args)

    exe_name = APP_NAME + (".exe" if os.name == "nt" else "")
    print(f"\nDone. Executable: {os.path.join(HERE, 'dist', exe_name)}")


if __name__ == "__main__":
    main()
