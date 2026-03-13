#!/usr/bin/env python3
"""LibreOffice CLI wrapper for PPT→PDF conversion.

Usage: python soffice.py --headless --convert-to pdf input.pptx
Handles macOS/Linux path differences and AF_UNIX sandbox limitations.
"""
import os
import subprocess
import sys


def find_soffice():
    """Find soffice binary across platforms."""
    candidates = [
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    from shutil import which
    found = which("soffice")
    if found:
        return found
    return None


def main():
    soffice = find_soffice()
    if not soffice:
        print("Error: LibreOffice (soffice) not found", file=sys.stderr)
        sys.exit(1)

    args = [soffice] + sys.argv[1:]
    profile_dir = os.path.join(os.path.expanduser("~"), ".claude", "soffice_profile")
    os.makedirs(profile_dir, exist_ok=True)
    args.insert(1, f"-env:UserInstallation=file://{profile_dir}")

    result = subprocess.run(args, capture_output=True, text=True, timeout=120)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
