"""Windows EXE 빌드 스크립트 (PyInstaller).

    pip install pyinstaller
    python build_exe.py

결과: dist/PPT_Photo_Macro_Generator_v<버전>.exe  (파이썬 없이 실행되는 단일 파일)
버전은 ppt_macro_gen/__init__.py 의 __version__ 에서 읽어 파일명과 EXE 속성(자세히 탭)에 넣습니다.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from ppt_macro_gen import APP_NAME, APP_NAME_ASCII, __version__  # noqa: E402


def version_tuple(v: str) -> tuple[int, int, int, int]:
    parts = [int(p) for p in v.split(".")]
    return tuple((parts + [0, 0, 0, 0])[:4])  # type: ignore[return-value]


def write_version_file(path: Path) -> None:
    t = version_tuple(__version__)
    path.write_text(
        f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={t}, prodvers={t}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041204B0', [
      StringStruct('FileDescription', {APP_NAME!r}),
      StringStruct('ProductName', {APP_NAME!r}),
      StringStruct('FileVersion', {__version__!r}),
      StringStruct('ProductVersion', {__version__!r}),
      StringStruct('InternalName', {APP_NAME_ASCII!r}),
      StringStruct('OriginalFilename', {APP_NAME_ASCII + '.exe'!r}),
    ])]),
    VarFileInfo([VarStruct('Translation', [0x0412, 1200])])
  ]
)
""",
        encoding="utf-8",
    )


def main() -> None:
    name = f"{APP_NAME_ASCII}_v{__version__}"
    build = ROOT / "build"
    build.mkdir(exist_ok=True)
    vfile = build / "version_info.txt"
    write_version_file(vfile)
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", name,
        "--specpath", str(build),
        "--workpath", str(build / "work"),
        "--distpath", str(ROOT / "dist"),
        str(ROOT / "run.py"),
    ]
    if sys.platform == "win32":
        cmd[-1:-1] = ["--version-file", str(vfile)]
    print(" ".join(cmd))
    subprocess.run(cmd, check=True, cwd=ROOT)
    print(f"\n완료: dist/{name}{'.exe' if sys.platform == 'win32' else ''}")


if __name__ == "__main__":
    main()
