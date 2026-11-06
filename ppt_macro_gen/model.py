"""격자 레이아웃 모델: 행/열 크기(cm)로부터 사진 칸(슬롯) 좌표를 계산한다.

좌표계: 기준 개체(표/도형)의 왼쪽 위 꼭짓점 = (0, 0), 단위 cm, 오른쪽/아래가 +.
"""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from typing import Iterable

from . import __version__

SCHEMA_VERSION = 1

# 사진 크기 결정 방식
SIZE_MARGIN = "margin"  # 칸 크기 - 여백 (칸마다 크기가 달라도 자동으로 맞춤)
SIZE_FIXED = "fixed"  # 모든 사진 고정 크기, 칸 정중앙

# 비율
ASPECT_STRETCH = "stretch"  # 사진 박스에 강제로 맞춤 (원본 비율이 다르면 찌그러짐)
ASPECT_FIT = "fit"  # 원본 비율 유지, 박스 안에 들어가도록 축소 후 중앙
ASPECT_FILL = "fill"  # 원본 비율 유지, 박스를 꽉 채우고 넘치는 부분은 잘라냄 (PPT 자르기)

# 사진 가져오기 방식
SOURCE_FILE_DIALOG = "file_dialog"  # 파일 선택 창에서 다중 선택해 새로 삽입
SOURCE_SELECTED = "selected"  # PPT에서 선택한 사진을 이동/크기 변경
SOURCE_FOLDER = "folder"  # 폴더에서 1.jpg ~ N.jpg 를 번호로 찾아 삽입

# 파일 정렬
SORT_NATURAL = "natural"  # 1, 2, ..., 10 (숫자 인식)
SORT_TEXT = "text"  # 1, 10, 2 (단순 문자열)

# 선택 사진 순서
ORDER_POSITION = "position"  # 기존 위치: 위→아래, 같은 줄은 왼쪽→오른쪽
ORDER_SELECTION = "selection"  # ShapeRange 순서 (Ctrl 클릭 순서와 다를 수 있음)

# 칸 채우기 순서
FILL_ROW = "row"  # 행 우선: 1 2 / 3 4
FILL_COL = "col"  # 열 우선: 1 3 / 2 4

VBA_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,59}$")


@dataclass
class Slot:
    """사진 1장이 들어갈 자리. 모든 값은 기준 개체 왼쪽 위 기준 cm."""

    index: int  # 1부터
    row: int  # 0부터
    col: int  # 0부터
    cell_x: float
    cell_y: float
    cell_w: float
    cell_h: float
    x: float  # 사진 박스 왼쪽 위
    y: float
    w: float  # 사진 박스 크기
    h: float
    shrunk: bool = False  # 고정 크기가 칸보다 커서 칸 − 여백으로 줄였는지


@dataclass
class Layout:
    macro_name: str = "InsertPhotos_Grid"
    col_widths: list[float] = field(default_factory=lambda: [6.7, 6.7])
    row_heights: list[float] = field(default_factory=lambda: [4.7, 4.7, 4.7])
    # 사진을 넣지 않는 칸 (텍스트 칸 등). [행, 열], 0부터
    excluded: list[list[int]] = field(default_factory=list)
    # 격자 시작점: 기준 개체 왼쪽 위에서 떨어진 거리
    origin_x: float = 0.0
    origin_y: float = 0.0
    fill_order: str = FILL_ROW
    size_mode: str = SIZE_MARGIN
    margin_x: float = 0.1  # 칸 왼쪽/오른쪽 각각
    margin_y: float = 0.1  # 칸 위/아래 각각
    pic_w: float = 6.5
    pic_h: float = 4.5
    # 고정 크기 모드에서 사진이 칸보다 크면 칸 − 여백으로 줄임 (끄면 칸 밖으로 튀어나옴)
    shrink_to_cell: bool = True
    aspect: str = ASPECT_STRETCH
    # 잘라서 채우기: 사진에서 시료(배경과 다른 부분)를 찾아 잘리지 않게 자를 위치를 정함
    smart_crop: bool = True
    border: bool = True
    border_weight: float = 0.75  # pt
    border_color: str = "#000000"
    source: str = SOURCE_FILE_DIALOG
    sort: str = SORT_NATURAL
    selected_order: str = ORDER_POSITION
    exact_count: bool = True
    extensions: list[str] = field(default_factory=lambda: ["jpg", "jpeg", "png", "bmp"])
    option_explicit: bool = True

    # ---- 격자 크기 ----
    @property
    def rows(self) -> int:
        return len(self.row_heights)

    @property
    def cols(self) -> int:
        return len(self.col_widths)

    def resize(self, rows: int, cols: int) -> None:
        """행/열 개수 변경. 새 칸은 마지막 크기를 복사하고 범위 밖 제외 칸은 지운다."""
        if rows < 1 or cols < 1:
            raise ValueError("행/열은 1 이상이어야 합니다.")
        self.row_heights = _resize_list(self.row_heights, rows, 4.7)
        self.col_widths = _resize_list(self.col_widths, cols, 6.7)
        self.excluded = [[r, c] for r, c in self.excluded if r < rows and c < cols]

    def is_excluded(self, row: int, col: int) -> bool:
        return [row, col] in self.excluded

    def toggle_cell(self, row: int, col: int) -> None:
        if self.is_excluded(row, col):
            self.excluded.remove([row, col])
        else:
            self.excluded.append([row, col])
            self.excluded.sort()

    @property
    def total_width(self) -> float:
        return sum(self.col_widths)

    @property
    def total_height(self) -> float:
        return sum(self.row_heights)

    # ---- 좌표 계산 ----
    def cell_origin(self, row: int, col: int) -> tuple[float, float]:
        return (
            self.origin_x + sum(self.col_widths[:col]),
            self.origin_y + sum(self.row_heights[:row]),
        )

    def cell_order(self) -> Iterable[tuple[int, int]]:
        if self.fill_order == FILL_COL:
            return ((r, c) for c in range(self.cols) for r in range(self.rows))
        return ((r, c) for r in range(self.rows) for c in range(self.cols))

    def slots(self) -> list[Slot]:
        result: list[Slot] = []
        for r, c in self.cell_order():
            if self.is_excluded(r, c):
                continue
            cx, cy = self.cell_origin(r, c)
            cw, ch = self.col_widths[c], self.row_heights[r]
            shrunk = False
            if self.size_mode == SIZE_FIXED:
                w, h = self.pic_w, self.pic_h
                # 칸 자체보다 클 때만 줄임 (칸 안에 들어가는 사진은 기존과 동일)
                if self.shrink_to_cell and (w > cw + 1e-9 or h > ch + 1e-9):
                    w, h, shrunk = min(w, cw - 2 * self.margin_x), min(h, ch - 2 * self.margin_y), True
            else:
                w, h = cw - 2 * self.margin_x, ch - 2 * self.margin_y
            x = cx + (cw - w) / 2
            y = cy + (ch - h) / 2
            result.append(
                Slot(
                    index=len(result) + 1,
                    row=r,
                    col=c,
                    cell_x=_r(cx),
                    cell_y=_r(cy),
                    cell_w=_r(cw),
                    cell_h=_r(ch),
                    x=_r(x),
                    y=_r(y),
                    w=_r(w),
                    h=_r(h),
                    shrunk=shrunk,
                )
            )
        return result

    # ---- 검증 ----
    def errors(self) -> list[str]:
        """코드를 생성할 수 없는 문제."""
        errs: list[str] = []
        if not VBA_IDENT.match(self.macro_name):
            errs.append("매크로 이름은 영문자로 시작하고 영문/숫자/_ 만 쓸 수 있습니다 (최대 60자).")
        if any(v <= 0 for v in self.col_widths):
            errs.append("열 너비는 0보다 커야 합니다.")
        if any(v <= 0 for v in self.row_heights):
            errs.append("행 높이는 0보다 커야 합니다.")
        if self.size_mode == SIZE_FIXED and (self.pic_w <= 0 or self.pic_h <= 0):
            errs.append("사진 크기는 0보다 커야 합니다.")
        if self.border and self.border_weight <= 0:
            errs.append("테두리 두께는 0보다 커야 합니다.")
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", self.border_color or ""):
            errs.append("테두리 색은 #RRGGBB 형식이어야 합니다.")
        if self.source == SOURCE_FOLDER and not self.extensions:
            errs.append("폴더 모드에는 확장자가 1개 이상 필요합니다.")
        if not errs:
            slots = self.slots()
            if not slots:
                errs.append("사진 칸이 하나도 없습니다 (모든 칸이 제외됨).")
            for s in slots:
                if s.w <= 0 or s.h <= 0:
                    errs.append(f"사진 {s.index}: 여백이 칸({s.cell_w}×{s.cell_h})보다 커서 사진 크기가 0 이하입니다.")
        return errs

    def warnings(self) -> list[str]:
        """생성은 가능하지만 확인이 필요한 점."""
        warns: list[str] = []
        if self.errors():
            return warns
        for s in self.slots():
            if s.shrunk:
                warns.append(f"사진 {s.index}: 고정 크기({self.pic_w}×{self.pic_h})가 칸({s.cell_w}×{s.cell_h})보다 커서 "
                             f"{s.w}×{s.h}로 줄였습니다.")
            elif s.w > s.cell_w + 1e-9 or s.h > s.cell_h + 1e-9:
                warns.append(f"사진 {s.index}: 사진({s.w}×{s.h})이 칸({s.cell_w}×{s.cell_h})보다 커서 옆 칸과 겹칩니다.")
        if self.source == SOURCE_SELECTED and self.selected_order == ORDER_SELECTION:
            warns.append("선택 순서(ShapeRange)는 Ctrl 클릭 순서와 다를 수 있습니다.")
        return warns

    # ---- 저장/불러오기 ----
    def to_dict(self) -> dict:
        d = asdict(self)
        d["schema"] = SCHEMA_VERSION
        d["generator_version"] = __version__
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Layout":
        known = {f for f in cls.__dataclass_fields__}
        data = {k: v for k, v in d.items() if k in known}
        lay = cls(**data)
        lay.col_widths = [float(v) for v in lay.col_widths]
        lay.row_heights = [float(v) for v in lay.row_heights]
        lay.excluded = sorted([int(r), int(c)] for r, c in lay.excluded)
        return lay

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, text: str) -> "Layout":
        return cls.from_dict(json.loads(text))


def _r(v: float) -> float:
    """부동소수점 오차 제거 (1.46 + 0.1 = 1.5599999 → 1.56)."""
    return round(v + 0.0, 6) + 0.0


def _resize_list(values: list[float], n: int, default: float) -> list[float]:
    if n <= len(values):
        return list(values[:n])
    fill = values[-1] if values else default
    return list(values) + [fill] * (n - len(values))


# ---- 기본 제공 프리셋 (인계 문서 기준) ----
# 사진 크기는 모두 '칸 크기 − 여백' 모드가 기본. 여백은 기존 고정 크기와 똑같은 결과가 나오도록 정했고,
# pic_w/pic_h 에는 기존 고정 크기를 남겨 두어 '고정 크기'로 바꾸면 바로 같은 값이 쓰이게 했다.
def presets() -> dict[str, Layout]:
    return {
        "최신: 파일 선택 6장, 왼쪽 1.46cm 열 + 2×3": Layout(
            macro_name="Insert6Photos_2x3_LeftColumn",
            col_widths=[1.46, 6.7, 6.7],
            row_heights=[4.7, 4.7, 4.7],
            excluded=[[0, 0], [1, 0], [2, 0]],
            size_mode=SIZE_MARGIN,
            margin_x=0.1,
            margin_y=0.1,
            pic_w=6.5,
            pic_h=4.5,
            aspect=ASPECT_FILL,
            source=SOURCE_FILE_DIALOG,
        ),
        "G: 선택 사진 6장 2×3 (6.4×4.4, 간격 0.1)": Layout(
            macro_name="Arrange6Selected_2x3",
            col_widths=[6.5, 6.5],
            row_heights=[4.5, 4.5, 4.5],
            origin_x=0.0,
            origin_y=0.74,
            size_mode=SIZE_MARGIN,
            margin_x=0.05,
            margin_y=0.05,
            pic_w=6.4,
            pic_h=4.4,
            border=False,
            source=SOURCE_SELECTED,
        ),
        "E: 선택 사진 15장 3×5 (4.7×2.98)": Layout(
            macro_name="Arrange15Selected_3x5",
            col_widths=[4.85] * 3,
            row_heights=[3.82] * 5,
            origin_x=-0.025,
            origin_y=0.37,
            size_mode=SIZE_MARGIN,
            margin_x=0.075,
            margin_y=0.42,
            pic_w=4.7,
            pic_h=2.98,
            border=False,
            source=SOURCE_SELECTED,
        ),
        "M: 선택 사진 1장 (9.9×6.9, +0.1/+0.1)": Layout(
            macro_name="Arrange1Selected",
            col_widths=[10.1],
            row_heights=[7.1],
            size_mode=SIZE_MARGIN,
            margin_x=0.1,
            margin_y=0.1,
            pic_w=9.9,
            pic_h=6.9,
            border=False,
            source=SOURCE_SELECTED,
        ),
    }
