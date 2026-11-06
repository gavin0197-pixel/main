"""tkinter GUI: 격자/칸 크기 입력 → 미리보기 → VBA 코드 복사/저장."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from . import APP_NAME, __version__
from .model import (
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
from .vba import generate, num

MAX_CELLS = 30  # 행/열 최대 개수
ENTRIES_PER_LINE = 8


def parse_cm(text: str, what: str) -> float:
    t = text.strip().replace(",", ".")
    try:
        return float(t)
    except ValueError:
        raise ValueError(f"{what}: 숫자가 아닙니다 ('{text}')") from None


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=6)
        self.master = master
        self.layout = list(presets().values())[0]
        self.current_file: Path | None = None
        self._suspend = False
        self._col_vars: list[tk.StringVar] = []
        self._row_vars: list[tk.StringVar] = []
        self._cell_rects: list[tuple[float, float, float, float, int, int]] = []

        self._make_vars()
        self._build()
        self.v_preset.set(list(presets().keys())[0])
        self._load_layout(self.layout)

    # ------------------------------------------------------------------ vars
    def _make_vars(self) -> None:
        V = tk.StringVar
        self.v_name = V()
        self.v_rows = V()
        self.v_cols = V()
        self.v_all_col = V(value="6.7")
        self.v_all_row = V(value="4.7")
        self.v_origin_x = V()
        self.v_origin_y = V()
        self.v_fill = V()
        self.v_size_mode = V()
        self.v_margin_x = V()
        self.v_margin_y = V()
        self.v_pic_w = V()
        self.v_pic_h = V()
        self.v_aspect = V()
        self.v_shrink = tk.BooleanVar()
        self.v_smart = tk.BooleanVar()
        self.v_border = tk.BooleanVar()
        self.v_border_w = V()
        self.v_border_color = V()
        self.v_source = V()
        self.v_sort = V()
        self.v_sel_order = V()
        self.v_exact = tk.BooleanVar()
        self.v_exts = V()
        self.v_opt_explicit = tk.BooleanVar()
        self.v_preset = V()
        self.v_status = V()
        for var in (
            self.v_name, self.v_origin_x, self.v_origin_y, self.v_fill, self.v_size_mode,
            self.v_margin_x, self.v_margin_y, self.v_pic_w, self.v_pic_h, self.v_aspect, self.v_shrink, self.v_smart,
            self.v_border, self.v_border_w, self.v_border_color, self.v_source, self.v_sort,
            self.v_sel_order, self.v_exact, self.v_exts, self.v_opt_explicit,
        ):
            var.trace_add("write", lambda *_: self._on_change())

    # ----------------------------------------------------------------- build
    def _build(self) -> None:
        self.master.title(f"{APP_NAME}  v{__version__}")
        self.grid(sticky="nsew")
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        # 상단 도구줄
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(bar, text="프리셋").pack(side="left")
        self.cb_preset = ttk.Combobox(bar, textvariable=self.v_preset, state="readonly",
                                      values=list(presets().keys()), width=42)
        self.cb_preset.pack(side="left", padx=4)
        ttk.Button(bar, text="불러오기", command=self._apply_preset).pack(side="left")
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="설정 열기…", command=self._open_json).pack(side="left")
        ttk.Button(bar, text="설정 저장…", command=self._save_json).pack(side="left", padx=4)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="코드 복사", command=self._copy_code).pack(side="left")
        ttk.Button(bar, text=".bas 저장…", command=self._save_bas).pack(side="left", padx=4)
        ttk.Label(bar, text=f"버전 v{__version__}", foreground="#555").pack(side="right", padx=(8, 2))

        # 왼쪽: 설정 탭
        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="ns", padx=(0, 6))
        self._build_grid_tab(nb)
        self._build_photo_tab(nb)
        self._build_source_tab(nb)

        # 오른쪽: 미리보기 + 코드
        right = ttk.PanedWindow(self, orient="vertical")
        right.grid(row=1, column=1, sticky="nsew")

        pv = ttk.LabelFrame(right, text="미리보기 (칸 클릭 → 사진 칸 / 제외 칸 전환)", padding=4)
        self.canvas = tk.Canvas(pv, background="white", highlightthickness=0, width=620, height=380)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._draw_preview())
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        right.add(pv, weight=3)

        cf = ttk.LabelFrame(right, text="생성된 VBA 코드", padding=4)
        self.code = tk.Text(cf, wrap="none", height=14, font=("Consolas", 10), undo=False)
        ys = ttk.Scrollbar(cf, orient="vertical", command=self.code.yview)
        xs = ttk.Scrollbar(cf, orient="horizontal", command=self.code.xview)
        self.code.configure(yscrollcommand=ys.set, xscrollcommand=xs.set, state="disabled")
        self.code.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)
        right.add(cf, weight=2)

        status = ttk.Label(self, textvariable=self.v_status, anchor="w", justify="left", wraplength=1000)
        status.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.status_label = status

    def _section(self, parent, text: str) -> ttk.LabelFrame:
        f = ttk.LabelFrame(parent, text=text, padding=6)
        f.pack(fill="x", pady=(0, 6))
        return f

    def _entry(self, parent, var, width=7) -> ttk.Entry:
        return ttk.Entry(parent, textvariable=var, width=width)

    def _build_grid_tab(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="격자")

        f = self._section(tab, "매크로")
        ttk.Label(f, text="Sub 이름").grid(row=0, column=0, sticky="w")
        self._entry(f, self.v_name, 34).grid(row=0, column=1, sticky="w", padx=4)

        f = self._section(tab, "행 / 열 개수")
        ttk.Label(f, text="행").grid(row=0, column=0)
        sp_r = ttk.Spinbox(f, from_=1, to=MAX_CELLS, textvariable=self.v_rows, width=5,
                           command=self._on_resize)
        sp_r.grid(row=0, column=1, padx=4)
        ttk.Label(f, text="열").grid(row=0, column=2)
        sp_c = ttk.Spinbox(f, from_=1, to=MAX_CELLS, textvariable=self.v_cols, width=5,
                           command=self._on_resize)
        sp_c.grid(row=0, column=3, padx=4)
        for sp in (sp_r, sp_c):
            sp.bind("<Return>", lambda e: self._on_resize())
            sp.bind("<FocusOut>", lambda e: self._on_resize())

        f = self._section(tab, "열 너비 (cm) — 왼쪽부터")
        self.col_frame = ttk.Frame(f)
        self.col_frame.pack(fill="x")
        g = ttk.Frame(f)
        g.pack(fill="x", pady=(6, 0))
        ttk.Label(g, text="모든 열을").pack(side="left")
        self._entry(g, self.v_all_col).pack(side="left", padx=4)
        ttk.Button(g, text="cm로", command=lambda: self._set_all(self._col_vars, self.v_all_col)).pack(side="left")

        f = self._section(tab, "행 높이 (cm) — 위부터")
        self.row_frame = ttk.Frame(f)
        self.row_frame.pack(fill="x")
        g = ttk.Frame(f)
        g.pack(fill="x", pady=(6, 0))
        ttk.Label(g, text="모든 행을").pack(side="left")
        self._entry(g, self.v_all_row).pack(side="left", padx=4)
        ttk.Button(g, text="cm로", command=lambda: self._set_all(self._row_vars, self.v_all_row)).pack(side="left")

        f = self._section(tab, "격자 시작점 (기준 개체 왼쪽 위에서, cm)")
        ttk.Label(f, text="+X").grid(row=0, column=0)
        self._entry(f, self.v_origin_x).grid(row=0, column=1, padx=4)
        ttk.Label(f, text="+Y").grid(row=0, column=2)
        self._entry(f, self.v_origin_y).grid(row=0, column=3, padx=4)

        f = self._section(tab, "사진 번호 순서")
        ttk.Radiobutton(f, text="행 우선 (1 2 / 3 4)", variable=self.v_fill, value=FILL_ROW).pack(anchor="w")
        ttk.Radiobutton(f, text="열 우선 (1 3 / 2 4)", variable=self.v_fill, value=FILL_COL).pack(anchor="w")

        ttk.Label(tab, text="※ 텍스트가 들어가는 칸은 미리보기에서 클릭해 '제외'로 바꾸세요.",
                  foreground="#555").pack(anchor="w")

    def _build_photo_tab(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="사진")

        f = self._section(tab, "사진 크기")
        ttk.Radiobutton(f, text="칸 크기 − 여백 (칸마다 크기가 달라도 자동)", variable=self.v_size_mode,
                        value=SIZE_MARGIN).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(f, text="좌우 여백 각").grid(row=1, column=0, sticky="e", padx=(18, 0))
        self._entry(f, self.v_margin_x).grid(row=1, column=1, padx=4)
        ttk.Label(f, text="상하 여백 각").grid(row=1, column=2, sticky="e")
        self._entry(f, self.v_margin_y).grid(row=1, column=3, padx=4)
        ttk.Radiobutton(f, text="고정 크기 (칸 정중앙)", variable=self.v_size_mode,
                        value=SIZE_FIXED).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Label(f, text="가로").grid(row=3, column=0, sticky="e", padx=(18, 0))
        self._entry(f, self.v_pic_w).grid(row=3, column=1, padx=4)
        ttk.Label(f, text="세로").grid(row=3, column=2, sticky="e")
        self._entry(f, self.v_pic_h).grid(row=3, column=3, padx=4)
        ttk.Checkbutton(f, text="칸보다 크면 칸 − 여백으로 줄이기 (텍스트 칸 등 좁은 칸)",
                        variable=self.v_shrink).grid(row=4, column=0, columnspan=4, sticky="w", padx=(18, 0), pady=(4, 0))

        f = self._section(tab, "원본 비율")
        ttk.Radiobutton(f, text="무시하고 박스에 맞춤 (기존 매크로 방식, 찌그러질 수 있음)",
                        variable=self.v_aspect, value=ASPECT_STRETCH).pack(anchor="w")
        ttk.Radiobutton(f, text="유지 (박스 안에 들어가게 축소, 중앙 정렬)",
                        variable=self.v_aspect, value=ASPECT_FIT).pack(anchor="w")
        ttk.Radiobutton(f, text="유지 + 잘라서 꽉 채우기 (넘치는 부분은 PPT 자르기)",
                        variable=self.v_aspect, value=ASPECT_FILL).pack(anchor="w")
        self.cb_smart = ttk.Checkbutton(
            f, text="시료 자동 감지: 흰 배경과 다른 부분을 찾아 시료가 잘리지 않게 자르기\n"
                    "(시료가 너무 크면 그 방향만 칸보다 좁게 / 선택 사진 모드는 가운데 기준)",
            variable=self.v_smart)
        self.cb_smart.pack(anchor="w", padx=(18, 0))

        f = self._section(tab, "테두리")
        ttk.Checkbutton(f, text="실선 테두리", variable=self.v_border).grid(row=0, column=0, sticky="w")
        ttk.Label(f, text="두께(pt)").grid(row=0, column=1, padx=(10, 0))
        self._entry(f, self.v_border_w, 5).grid(row=0, column=2, padx=4)
        ttk.Label(f, text="색").grid(row=0, column=3)
        self._entry(f, self.v_border_color, 8).grid(row=0, column=4, padx=4)
        ttk.Button(f, text="…", width=3, command=self._pick_color).grid(row=0, column=5)

    def _build_source_tab(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="가져오기")

        f = self._section(tab, "사진 가져오기 방식")
        ttk.Radiobutton(f, text="파일 선택 창에서 여러 장 선택 → 새로 삽입", variable=self.v_source,
                        value=SOURCE_FILE_DIALOG).pack(anchor="w")
        ttk.Radiobutton(f, text="PPT에서 선택한 사진 → 위치/크기만 변경", variable=self.v_source,
                        value=SOURCE_SELECTED).pack(anchor="w")
        ttk.Radiobutton(f, text="폴더에서 1, 2, 3… 번호 파일 → 새로 삽입", variable=self.v_source,
                        value=SOURCE_FOLDER).pack(anchor="w")

        f = self._section(tab, "파일 선택 창: 파일명 정렬")
        self.rb_sort = [
            ttk.Radiobutton(f, text="숫자 인식 (1, 2, … 10)", variable=self.v_sort, value=SORT_NATURAL),
            ttk.Radiobutton(f, text="문자열 순 (1, 10, 2)", variable=self.v_sort, value=SORT_TEXT),
        ]
        for rb in self.rb_sort:
            rb.pack(anchor="w")

        f = self._section(tab, "선택 사진: 번호 순서")
        self.rb_order = [
            ttk.Radiobutton(f, text="기존 위치 (위→아래, 왼쪽→오른쪽)", variable=self.v_sel_order,
                            value=ORDER_POSITION),
            ttk.Radiobutton(f, text="선택 순서 (Ctrl 클릭 순서와 다를 수 있음)", variable=self.v_sel_order,
                            value=ORDER_SELECTION),
        ]
        for rb in self.rb_order:
            rb.pack(anchor="w")

        f = self._section(tab, "기타")
        ttk.Checkbutton(f, text="사진 수가 칸 수와 정확히 같아야 함\n(끄면 적은 장수도 앞 칸부터 채움)",
                        variable=self.v_exact).pack(anchor="w")
        g = ttk.Frame(f)
        g.pack(anchor="w", pady=(6, 0))
        ttk.Label(g, text="확장자").pack(side="left")
        self._entry(g, self.v_exts, 24).pack(side="left", padx=4)
        ttk.Checkbutton(f, text="맨 위에 Option Explicit 넣기 (새 모듈에 붙여넣을 때)",
                        variable=self.v_opt_explicit).pack(anchor="w", pady=(6, 0))

    # -------------------------------------------------------- size entries
    def _rebuild_size_entries(self) -> None:
        for frame, values, attr, label in (
            (self.col_frame, self.layout.col_widths, "_col_vars", "열"),
            (self.row_frame, self.layout.row_heights, "_row_vars", "행"),
        ):
            for w in frame.winfo_children():
                w.destroy()
            vars_: list[tk.StringVar] = []
            for i, v in enumerate(values):
                sv = tk.StringVar(value=num(v))
                sv.trace_add("write", lambda *_: self._on_change())
                cell = ttk.Frame(frame)
                cell.grid(row=i // ENTRIES_PER_LINE, column=i % ENTRIES_PER_LINE, padx=2, pady=2)
                ttk.Label(cell, text=f"{i + 1}{label}", foreground="#555").pack()
                ttk.Entry(cell, textvariable=sv, width=6, justify="center").pack()
                vars_.append(sv)
            setattr(self, attr, vars_)

    def _set_all(self, vars_: list[tk.StringVar], src: tk.StringVar) -> None:
        self._suspend = True
        for v in vars_:
            v.set(src.get())
        self._suspend = False
        self._on_change()

    def _on_resize(self) -> None:
        try:
            rows = int(self.v_rows.get())
            cols = int(self.v_cols.get())
        except ValueError:
            return
        rows = max(1, min(MAX_CELLS, rows))
        cols = max(1, min(MAX_CELLS, cols))
        if rows == self.layout.rows and cols == self.layout.cols:
            return
        try:
            self._read_sizes()
        except ValueError:
            pass
        self.layout.resize(rows, cols)
        self._rebuild_size_entries()
        self._on_change()

    # ------------------------------------------------------- model <-> ui
    def _load_layout(self, lay: Layout) -> None:
        self._suspend = True
        self.layout = lay
        self.v_name.set(lay.macro_name)
        self.v_rows.set(str(lay.rows))
        self.v_cols.set(str(lay.cols))
        self.v_origin_x.set(num(lay.origin_x))
        self.v_origin_y.set(num(lay.origin_y))
        self.v_fill.set(lay.fill_order)
        self.v_size_mode.set(lay.size_mode)
        self.v_margin_x.set(num(lay.margin_x))
        self.v_margin_y.set(num(lay.margin_y))
        self.v_pic_w.set(num(lay.pic_w))
        self.v_pic_h.set(num(lay.pic_h))
        self.v_shrink.set(lay.shrink_to_cell)
        self.v_smart.set(lay.smart_crop)
        self.v_aspect.set(lay.aspect)
        self.v_border.set(lay.border)
        self.v_border_w.set(num(lay.border_weight))
        self.v_border_color.set(lay.border_color)
        self.v_source.set(lay.source)
        self.v_sort.set(lay.sort)
        self.v_sel_order.set(lay.selected_order)
        self.v_exact.set(lay.exact_count)
        self.v_exts.set(", ".join(lay.extensions))
        self.v_opt_explicit.set(lay.option_explicit)
        self._rebuild_size_entries()
        self._suspend = False
        self._on_change()

    def _read_sizes(self) -> None:
        self.layout.col_widths = [parse_cm(v.get(), f"{i + 1}열 너비") for i, v in enumerate(self._col_vars)]
        self.layout.row_heights = [parse_cm(v.get(), f"{i + 1}행 높이") for i, v in enumerate(self._row_vars)]

    def _read_ui(self) -> None:
        """UI 값을 self.layout 에 반영. 숫자 오류는 ValueError."""
        lay = self.layout
        self._read_sizes()
        lay.macro_name = self.v_name.get().strip()
        lay.origin_x = parse_cm(self.v_origin_x.get(), "시작점 X")
        lay.origin_y = parse_cm(self.v_origin_y.get(), "시작점 Y")
        lay.fill_order = self.v_fill.get()
        lay.size_mode = self.v_size_mode.get()
        lay.margin_x = parse_cm(self.v_margin_x.get(), "좌우 여백")
        lay.margin_y = parse_cm(self.v_margin_y.get(), "상하 여백")
        lay.pic_w = parse_cm(self.v_pic_w.get(), "사진 가로")
        lay.pic_h = parse_cm(self.v_pic_h.get(), "사진 세로")
        lay.shrink_to_cell = bool(self.v_shrink.get())
        lay.smart_crop = bool(self.v_smart.get())
        lay.aspect = self.v_aspect.get()
        lay.border = bool(self.v_border.get())
        lay.border_weight = parse_cm(self.v_border_w.get(), "테두리 두께")
        lay.border_color = self.v_border_color.get().strip()
        lay.source = self.v_source.get()
        lay.sort = self.v_sort.get()
        lay.selected_order = self.v_sel_order.get()
        lay.exact_count = bool(self.v_exact.get())
        lay.extensions = [e.strip().lstrip(".").lower() for e in self.v_exts.get().replace(";", ",").split(",")
                          if e.strip()]
        lay.option_explicit = bool(self.v_opt_explicit.get())

    def _on_change(self) -> None:
        if self._suspend:
            return
        self._update_enabled()
        try:
            self._read_ui()
        except ValueError as e:
            self._show_status([str(e)], [])
            self._set_code("")
            self._draw_preview(valid=False)
            return
        errs = self.layout.errors()
        warns = self.layout.warnings()
        self._show_status(errs, warns)
        self._set_code("" if errs else generate(self.layout).replace("\r\n", "\n"))
        self._draw_preview(valid=not errs)

    def _update_enabled(self) -> None:
        src = self.v_source.get()
        for rb in self.rb_sort:
            rb.state(["!disabled"] if src == SOURCE_FILE_DIALOG else ["disabled"])
        for rb in self.rb_order:
            rb.state(["!disabled"] if src == SOURCE_SELECTED else ["disabled"])
        smart_ok = self.v_aspect.get() == ASPECT_FILL and src != SOURCE_SELECTED
        self.cb_smart.state(["!disabled"] if smart_ok else ["disabled"])

    def _show_status(self, errs: list[str], warns: list[str]) -> None:
        lay = self.layout
        parts = []
        try:
            n = len(lay.slots())
            parts.append(f"격자 전체 {num(lay.total_width)} × {num(lay.total_height)} cm  |  사진 칸 {n}개")
        except Exception:
            pass
        parts += ["❌ " + e for e in errs]
        parts += ["⚠ " + w for w in warns]
        self.v_status.set("\n".join(parts))
        self.status_label.configure(foreground="#b00020" if errs else ("#8a5a00" if warns else "#1b5e20"))

    def _set_code(self, text: str) -> None:
        self.code.configure(state="normal")
        self.code.delete("1.0", "end")
        self.code.insert("1.0", text)
        self.code.configure(state="disabled")

    # ------------------------------------------------------------ preview
    def _draw_preview(self, valid: bool = True) -> None:
        c = self.canvas
        c.delete("all")
        self._cell_rects = []
        lay = self.layout
        W, H = c.winfo_width(), c.winfo_height()
        if W < 20 or H < 20 or not lay.col_widths or not lay.row_heights:
            return
        if any(v <= 0 for v in lay.col_widths + lay.row_heights):
            c.create_text(W / 2, H / 2, text="칸 크기를 확인하세요", fill="#b00020")
            return

        pad = 44
        min_x = min(0.0, lay.origin_x)
        min_y = min(0.0, lay.origin_y)
        max_x = lay.origin_x + lay.total_width
        max_y = lay.origin_y + lay.total_height
        span_x = max(max_x - min_x, 0.1)
        span_y = max(max_y - min_y, 0.1)
        s = min((W - 2 * pad) / span_x, (H - 2 * pad) / span_y)
        ox = pad - min_x * s
        oy = pad - min_y * s

        def px(x: float) -> float:
            return ox + x * s

        def py(y: float) -> float:
            return oy + y * s

        # 칸
        for r in range(lay.rows):
            for col in range(lay.cols):
                cx, cy = lay.cell_origin(r, col)
                x0, y0 = px(cx), py(cy)
                x1, y1 = px(cx + lay.col_widths[col]), py(cy + lay.row_heights[r])
                ex = lay.is_excluded(r, col)
                c.create_rectangle(x0, y0, x1, y1, outline="#888",
                                   fill="#eeeeee" if ex else "white", stipple="gray25" if ex else "")
                if ex:
                    c.create_text((x0 + x1) / 2, (y0 + y1) / 2, text="제외", fill="#777")
                self._cell_rects.append((x0, y0, x1, y1, r, col))

        # 사진 박스
        if valid:
            for sl in lay.slots():
                x0, y0 = px(sl.x), py(sl.y)
                x1, y1 = px(sl.x + sl.w), py(sl.y + sl.h)
                overflow = sl.w > sl.cell_w + 1e-9 or sl.h > sl.cell_h + 1e-9
                # 파랑: 정상 / 주황: 칸에 맞춰 줄어듦 / 빨강: 칸 밖으로 튀어나옴
                fill = "#ffcdd2" if overflow else ("#ffe0b2" if sl.shrunk else "#d6e4f5")
                c.create_rectangle(x0, y0, x1, y1, outline="black" if lay.border else "#3b6fb6",
                                   width=2 if lay.border else 1, fill=fill)
                c.create_text((x0 + x1) / 2, (y0 + y1) / 2 - 8, text=str(sl.index),
                              font=("", 14, "bold"), fill="#1f3f73")
                if x1 - x0 >= 50:  # 좁은 칸에서는 크기 글자 생략
                    c.create_text((x0 + x1) / 2, (y0 + y1) / 2 + 10,
                                  text=f"{num(sl.w)}×{num(sl.h)}", fill="#1f3f73", font=("", 8))

        # 열 너비 / 행 높이 표시
        for col in range(lay.cols):
            cx, _ = lay.cell_origin(0, col)
            c.create_text(px(cx + lay.col_widths[col] / 2), py(lay.origin_y) - 10,
                          text=num(lay.col_widths[col]), fill="#555", font=("", 8))
        for r in range(lay.rows):
            _, cy = lay.cell_origin(r, 0)
            c.create_text(px(lay.origin_x) - 14, py(cy + lay.row_heights[r] / 2),
                          text=num(lay.row_heights[r]), fill="#555", font=("", 8))

        # 기준 개체 왼쪽 위 (0,0)
        x0, y0 = px(0), py(0)
        c.create_line(x0 - 8, y0, x0 + 8, y0, fill="#d32f2f", width=2)
        c.create_line(x0, y0 - 8, x0, y0 + 8, fill="#d32f2f", width=2)
        c.create_text(x0 - 6, y0 - 4, text="기준", anchor="se", fill="#d32f2f", font=("", 8))

    def _on_canvas_click(self, event) -> None:
        for x0, y0, x1, y1, r, col in self._cell_rects:
            if x0 <= event.x <= x1 and y0 <= event.y <= y1:
                self.layout.toggle_cell(r, col)
                self._on_change()
                return

    # ------------------------------------------------------------ actions
    def _apply_preset(self) -> None:
        name = self.v_preset.get()
        if name:
            self._load_layout(presets()[name])

    def _pick_color(self) -> None:
        _, hexv = colorchooser.askcolor(color=self.v_border_color.get() or "#000000", parent=self)
        if hexv:
            self.v_border_color.set(hexv.upper())

    def _current_code(self, module_header: bool = False) -> str | None:
        try:
            self._read_ui()
            return generate(self.layout, module_header=module_header)
        except ValueError as e:
            messagebox.showerror("코드 생성 불가", str(e), parent=self)
            return None

    def _copy_code(self) -> None:
        code = self._current_code()
        if code is None:
            return
        self.clipboard_clear()
        self.clipboard_append(code)
        self.v_status.set(self.v_status.get().split("\n")[0] + "\n✅ 코드를 클립보드에 복사했습니다. VBA 편집기(Alt+F11) 모듈에 붙여넣으세요.")

    def _save_bas(self) -> None:
        code = self._current_code(module_header=True)
        if code is None:
            return
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".bas",
                                            initialfile=f"{self.layout.macro_name}.bas",
                                            filetypes=[("VBA 모듈", "*.bas")])
        if not path:
            return
        # VBA 편집기 가져오기(Import)는 시스템 ANSI 코드페이지를 사용 → 한국어 Windows는 cp949
        Path(path).write_bytes(code.encode("cp949", errors="replace"))
        messagebox.showinfo("저장 완료", f"{path}\n\nVBA 편집기 → 파일 → 파일 가져오기로 불러오세요.", parent=self)

    def _save_json(self) -> None:
        try:
            self._read_ui()
        except ValueError as e:
            messagebox.showerror("저장 불가", str(e), parent=self)
            return
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".json",
                                            initialfile=f"{self.layout.macro_name}.json",
                                            filetypes=[("배치 설정", "*.json")])
        if path:
            Path(path).write_text(self.layout.to_json(), encoding="utf-8")

    def _open_json(self) -> None:
        path = filedialog.askopenfilename(parent=self, filetypes=[("배치 설정", "*.json")])
        if not path:
            return
        try:
            lay = Layout.from_json(Path(path).read_text(encoding="utf-8"))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("열기 실패", str(e), parent=self)
            return
        self._load_layout(lay)


def main() -> None:
    root = tk.Tk()
    root.geometry("1180x780")
    root.minsize(980, 640)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
