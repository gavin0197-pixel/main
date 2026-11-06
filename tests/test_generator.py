import itertools
import re

import pytest

from ppt_macro_gen.model import (
    ASPECT_FILL,
    ASPECT_FIT,
    ASPECT_STRETCH,
    FILL_COL,
    FILL_ROW,
    ORDER_POSITION,
    ORDER_SELECTION,
    SIZE_FIXED,
    SIZE_MARGIN,
    SORT_NATURAL,
    SORT_TEXT,
    SOURCE_FILE_DIALOG,
    SOURCE_FOLDER,
    SOURCE_SELECTED,
    Layout,
    presets,
)
from ppt_macro_gen.vba import generate, num

P = presets()


def xy(lay):
    return [(s.x, s.y, s.w, s.h) for s in lay.slots()]


def test_latest_preset_matches_handover_table():
    lay = P["최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3"]
    assert xy(lay) == [
        (1.56, 0.1, 6.5, 4.5),
        (8.26, 0.1, 6.5, 4.5),
        (1.56, 4.8, 6.5, 4.5),
        (8.26, 4.8, 6.5, 4.5),
        (1.56, 9.5, 6.5, 4.5),
        (8.26, 9.5, 6.5, 4.5),
    ]
    # 여백 모드로 바꿔도 같은 결과 (6.7-0.2, 4.7-0.2)
    lay.size_mode = SIZE_MARGIN
    assert xy(lay)[0] == (1.56, 0.1, 6.5, 4.5)


def test_g_preset_pitch_and_start():
    s = P["G: 선택 사진 6장 2×3 (6.4×4.4, 간격 0.1)"].slots()
    assert (s[0].x, s[0].y) == (0.05, 0.79)
    assert pytest.approx(s[1].x - s[0].x) == 6.5
    assert pytest.approx(s[2].y - s[0].y) == 4.5
    assert (s[0].w, s[0].h) == (6.4, 4.4)


def test_e_preset_pitch_and_start():
    s = P["E: 선택 사진 15장 3×5 (4.7×2.98)"].slots()
    assert len(s) == 15
    assert (s[0].x, s[0].y) == (0.05, 0.79)
    assert pytest.approx(s[1].x - s[0].x) == 4.85
    assert pytest.approx(s[3].y - s[0].y) == 3.82


def test_m_preset():
    s = P["M: 선택 사진 1장 (9.9×6.9, +0.1/+0.1)"].slots()
    assert xy_one(s[0]) == (0.1, 0.1, 9.9, 6.9)


def xy_one(s):
    return (s.x, s.y, s.w, s.h)


def test_irregular_cells_margin_mode():
    lay = Layout(col_widths=[2.0, 5.0, 3.0], row_heights=[1.0, 4.0], excluded=[[0, 0], [0, 1], [0, 2], [1, 0]],
                 size_mode=SIZE_MARGIN, margin_x=0.1, margin_y=0.2)
    assert xy(lay) == [(2.1, 1.2, 4.8, 3.6), (7.1, 1.2, 2.8, 3.6)]


def test_column_major_order_and_origin():
    lay = Layout(col_widths=[1, 1], row_heights=[1, 1], fill_order=FILL_COL, origin_x=0.5, origin_y=-0.5,
                 margin_x=0, margin_y=0)
    assert [(s.row, s.col) for s in lay.slots()] == [(0, 0), (1, 0), (0, 1), (1, 1)]
    assert lay.slots()[0].x == 0.5 and lay.slots()[0].y == -0.5


def test_resize_keeps_values_and_drops_out_of_range_exclusions():
    lay = Layout(col_widths=[1.46, 6.7], row_heights=[4.7], excluded=[[0, 1]])
    lay.resize(2, 3)
    assert lay.col_widths == [1.46, 6.7, 6.7] and lay.row_heights == [4.7, 4.7]
    lay.resize(1, 1)
    assert lay.excluded == []


def test_errors():
    assert Layout(macro_name="1abc").errors()
    assert Layout(macro_name="사진배치").errors()
    assert Layout(col_widths=[0.1], margin_x=0.1).errors()  # 사진 폭 0 이하
    lay = Layout(col_widths=[1], row_heights=[1], excluded=[[0, 0]])
    assert any("하나도" in e for e in lay.errors())
    with pytest.raises(ValueError):
        generate(Layout(macro_name="bad name"))


def test_warning_when_fixed_photo_bigger_than_cell():
    lay = Layout(col_widths=[5], row_heights=[5], size_mode=SIZE_FIXED, pic_w=6, pic_h=4)
    assert lay.warnings()


def test_json_roundtrip():
    lay = P["최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3"]
    back = Layout.from_json(lay.to_json())
    assert back == lay


def test_num_format():
    assert num(1.56) == "1.56"
    assert num(4.0) == "4"
    assert num(-0.0) == "0"
    assert num(-0.025) == "-0.025"


# ---------------- VBA 구조 검사 ----------------

BLOCK_OPEN = [
    (re.compile(r"^(Private\s+)?Sub\s+\w+\("), "Sub"),
    (re.compile(r"^(Private\s+)?Function\s+\w+\("), "Function"),
    (re.compile(r"^For\b"), "For"),
    (re.compile(r"^Do\b(?!.*:\s*Loop\s*$)"), "Do"),
    (re.compile(r"^With\b"), "With"),
    (re.compile(r"^If\b.*\bThen\s*$"), "If"),
]
BLOCK_CLOSE = [
    (re.compile(r"^End Sub\b"), "Sub"),
    (re.compile(r"^End Function\b"), "Function"),
    (re.compile(r"^Next\b"), "For"),
    (re.compile(r"^Loop\b"), "Do"),
    (re.compile(r"^End With\b"), "With"),
    (re.compile(r"^End If\b"), "If"),
]


def check_blocks(code: str) -> None:
    stack = []
    for n, raw in enumerate(code.split("\r\n"), 1):
        line = raw.strip()
        if not line or line.startswith("'"):
            continue
        for rx, kind in BLOCK_OPEN:
            if rx.search(line):
                stack.append((kind, n))
                break
        else:
            for rx, kind in BLOCK_CLOSE:
                if rx.search(line):
                    assert stack and stack[-1][0] == kind, f"line {n}: {line!r} closes {stack[-1:]}"
                    stack.pop()
                    break
    assert not stack, f"unclosed blocks: {stack}"


def all_combos():
    for src, (aspect, smart), size, border, exact, sort, order, fill in itertools.product(
        [SOURCE_FILE_DIALOG, SOURCE_SELECTED, SOURCE_FOLDER],
        [(ASPECT_STRETCH, True), (ASPECT_FIT, True), (ASPECT_FILL, True), (ASPECT_FILL, False)],
        [SIZE_MARGIN, SIZE_FIXED],
        [True, False],
        [True, False],
        [SORT_NATURAL, SORT_TEXT],
        [ORDER_POSITION, ORDER_SELECTION],
        [FILL_ROW, FILL_COL],
    ):
        yield Layout(macro_name="TestMacro", source=src, aspect=aspect, smart_crop=smart, size_mode=size, border=border,
                     exact_count=exact, sort=sort, selected_order=order, fill_order=fill)


def test_generated_vba_is_structurally_valid_for_all_options():
    for lay in all_combos():
        code = generate(lay)
        check_blocks(code)
        assert code.count("Sub TestMacro()") == 1
        # 사용하는 보조 함수는 반드시 정의되어 있어야 함
        for helper in set(re.findall(r"\b(\w+_TestMacro)\b", code)):
            assert re.search(rf"(Sub|Function) {helper}\(", code), helper
        # 정의만 되고 안 쓰는 보조 함수도 없어야 함 (자기 본문 밖에서 호출되는지)
        for helper in re.findall(r"Private (?:Sub|Function) (\w+)\(", code):
            m = re.search(rf"Private (?:Sub|Function) {helper}\(.*?End (?:Sub|Function)", code, re.S)
            outside = code[: m.start()] + code[m.end():]
            assert re.search(rf"\b{helper}\b", outside), f"unused helper {helper}"
        # 회전된(휴대폰) 사진 처리
        assert "sideways = " in code
        assert ("ObjBox_TestMacro" in code) == (lay.aspect == ASPECT_FILL and lay.smart_crop
                                                and lay.source != SOURCE_SELECTED)
        assert "\n" not in code.replace("\r\n", "")  # CRLF 일관
        code.encode("cp949")  # .bas 저장 가능


def test_module_header_name_differs_from_sub():
    code = generate(Layout(macro_name="Foo"), module_header=True)
    assert code.startswith('Attribute VB_Name = "mod_Foo"')


def test_slot_lines_in_code():
    code = generate(P["최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3"])
    assert "slotX(2) = 8.26: slotY(2) = 0.1: slotW(2) = 6.5: slotH(2) = 4.5" in code
    assert "Const SLOT_COUNT As Long = 6" in code


# ---------------- 제외 칸 해제 시 사진이 칸 밖으로 튀어나오던 문제 ----------------

def _latest_with_2cm_text_col():
    lay = presets()["최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3"]
    lay.col_widths[0] = 2.0
    return lay


def test_toggle_twice_restores_identical_slots():
    lay = _latest_with_2cm_text_col()
    before = xy(lay)
    lay.toggle_cell(2, 0)
    lay.toggle_cell(2, 0)
    assert xy(lay) == before


def test_reenabled_narrow_cell_shrinks_fixed_photo_inside_cell():
    lay = _latest_with_2cm_text_col()
    lay.toggle_cell(2, 0)  # 3행 1열(텍스트 칸)을 사진 칸으로
    s = next(s for s in lay.slots() if (s.row, s.col) == (2, 0))
    assert s.shrunk and (s.x, s.y, s.w, s.h) == (0.1, 9.5, 1.8, 4.5)
    assert s.x >= s.cell_x and s.x + s.w <= s.cell_x + s.cell_w
    # 다른 칸은 그대로 6.5×4.5
    assert all((t.w, t.h) == (6.5, 4.5) for t in lay.slots() if t is not s and not t.shrunk)
    assert any("줄였습니다" in w for w in lay.warnings())
    assert "(칸에 맞춰 축소)" in generate(lay)


def test_shrink_off_keeps_old_overflow_behavior_with_warning():
    lay = _latest_with_2cm_text_col()
    lay.shrink_to_cell = False
    lay.toggle_cell(2, 0)
    s = next(s for s in lay.slots() if (s.row, s.col) == (2, 0))
    assert (s.x, s.w) == (-2.25, 6.5) and not s.shrunk
    assert any("겹칩니다" in w for w in lay.warnings())


def test_shrink_does_not_touch_photos_that_fit():
    for name, lay in presets().items():
        off = Layout.from_dict({**lay.to_dict(), "shrink_to_cell": False})
        assert xy(lay) == xy(off), name


def test_old_json_without_shrink_field_loads():
    d = P["최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3"].to_dict()
    d.pop("shrink_to_cell")
    assert Layout.from_dict(d).shrink_to_cell is True
