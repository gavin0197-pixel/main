import re
from pathlib import Path

from ppt_macro_gen import __version__
from ppt_macro_gen.model import Layout
from ppt_macro_gen.vba import generate

ROOT = Path(__file__).resolve().parent.parent


def test_version_format():
    assert re.fullmatch(r"\d+\.\d+\.\d+", __version__)


def test_changelog_has_current_version_on_top():
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    first = re.search(r"^## \[([^\]]+)\]", text, re.M)
    assert first and first.group(1) == __version__, "CHANGELOG.md 맨 위 섹션을 현재 버전으로 추가하세요"


def test_version_stamped_in_outputs():
    assert f"v{__version__}" in generate(Layout())
    assert Layout().to_dict()["generator_version"] == __version__
    assert "generator_version" not in Layout.from_dict(Layout().to_dict()).__dict__
