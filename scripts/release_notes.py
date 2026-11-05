"""CHANGELOG.md 에서 해당 버전 섹션만 출력: python scripts/release_notes.py 0.2.0"""
import re
import sys
from pathlib import Path

ver = sys.argv[1]
text = Path(__file__).resolve().parent.parent.joinpath("CHANGELOG.md").read_text(encoding="utf-8")
m = re.search(rf"^## \[{re.escape(ver)}\][^\n]*\n(.*?)(?=^## \[|\Z)", text, re.S | re.M)
sys.stdout.reconfigure(encoding="utf-8")
print(m.group(1).strip() if m else f"v{ver}")
