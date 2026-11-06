"""Layout → 실행 가능한 PowerPoint VBA 코드 생성."""

from __future__ import annotations

from . import APP_NAME, __version__
from .model import (
    ASPECT_FIT,
    FILL_COL,
    ORDER_POSITION,
    SIZE_FIXED,
    SORT_NATURAL,
    SOURCE_FILE_DIALOG,
    SOURCE_FOLDER,
    SOURCE_SELECTED,
    Layout,
)

SOURCE_LABEL = {
    SOURCE_FILE_DIALOG: "파일 선택 창에서 다중 선택 → 새로 삽입",
    SOURCE_SELECTED: "PPT에서 선택한 사진 → 이동/크기 변경",
    SOURCE_FOLDER: "폴더에서 1, 2, 3... 번호 파일 → 새로 삽입",
}


def num(v: float) -> str:
    """VBA 숫자 리터럴 (불필요한 0 제거)."""
    s = f"{v:.4f}".rstrip("0").rstrip(".")
    if s in ("-0", ""):
        s = "0"
    return s


def vba_str(s: str) -> str:
    return '"' + s.replace('"', '""') + '"'


def generate(lay: Layout, module_header: bool = False) -> str:
    errs = lay.errors()
    if errs:
        raise ValueError("\n".join(errs))

    name = lay.macro_name
    slots = lay.slots()
    n = len(slots)
    fit = lay.aspect == ASPECT_FIT
    color = lay.border_color.lstrip("#")
    rgb = f"RGB({int(color[0:2], 16)}, {int(color[2:4], 16)}, {int(color[4:6], 16)})"

    L: list[str] = []
    a = L.append

    if module_header:
        a(f'Attribute VB_Name = "mod_{name}"')
    if lay.option_explicit:
        a("Option Explicit")
        a("")

    # ---- 머리 주석: 설정 요약 ----
    a("' " + "=" * 70)
    a(f"' {name}")
    a(f"' {APP_NAME} v{__version__} 로 생성")
    a("' " + "-" * 70)
    a(f"' 사용법 : {_usage(lay, n)}")
    a(f"' 가져오기: {SOURCE_LABEL[lay.source]}")
    a(f"' 격자   : {lay.rows}행 × {lay.cols}열, 시작점 기준 개체 왼쪽 위 +X {num(lay.origin_x)} / +Y {num(lay.origin_y)} cm")
    a(f"' 열 너비 : {', '.join(num(v) for v in lay.col_widths)} cm")
    a(f"' 행 높이 : {', '.join(num(v) for v in lay.row_heights)} cm")
    if lay.excluded:
        a("' 제외 칸 : " + ", ".join(f"({r + 1}행 {c + 1}열)" for r, c in lay.excluded))
    if lay.size_mode == SIZE_FIXED:
        a(f"' 사진   : 가로 {num(lay.pic_w)} × 세로 {num(lay.pic_h)} cm 고정, 칸 정중앙"
          + (f" (칸보다 크면 칸 - 여백 {num(lay.margin_x)}/{num(lay.margin_y)} cm로 축소)" if lay.shrink_to_cell else ""))
    else:
        a(f"' 사진   : 칸 크기 - 여백 (좌우 각 {num(lay.margin_x)} / 상하 각 {num(lay.margin_y)} cm)")
    a("' 비율   : " + ("원본 비율 유지 (박스 안에 맞춤, 중앙)" if fit else "박스에 강제로 맞춤 (원본 비율이 다르면 찌그러짐)"))
    a("' 순서   : " + ("열 우선 (위→아래, 다음 열)" if lay.fill_order == FILL_COL else "행 우선 (왼쪽→오른쪽, 다음 행)"))
    a("' 테두리 : " + (f"실선 {num(lay.border_weight)}pt, {lay.border_color.upper()}" if lay.border else "없음"))
    a("' " + "=" * 70)
    a(f"Sub {name}()")
    a("")
    a(f"    Const SLOT_COUNT As Long = {n}")
    a("    Const CM As Double = 72 / 2.54   ' cm → pt")
    a("")
    a("    Dim sel As Selection")
    a("    Dim sld As Slide")
    a("    Dim refShp As Shape")
    a("    Dim pic As Shape")
    a("    Dim i As Long, j As Long, cnt As Long")
    a("    Dim baseX As Double, baseY As Double")
    a("    Dim slotX(1 To SLOT_COUNT) As Double, slotY(1 To SLOT_COUNT) As Double")
    a("    Dim slotW(1 To SLOT_COUNT) As Double, slotH(1 To SLOT_COUNT) As Double")
    if lay.source == SOURCE_SELECTED:
        a("    Dim pics() As Shape")
        a("    Dim shp As Shape")
        a("    Dim tmpShp As Shape")
        a("    Dim refCount As Long")
    else:
        a("    Dim files(1 To SLOT_COUNT) As String")
        if lay.source == SOURCE_FILE_DIALOG:
            a("    Dim fd As FileDialog")
            a("    Dim tmp As String")
        else:
            a("    Dim fd As FileDialog")
            a("    Dim folder As String")
            a("    Dim exts As Variant")
            a("    Dim k As Long")
            a("    Dim missing As String")
    a("")

    # ---- 슬롯 표 ----
    a("    ' 사진 박스 (cm, 기준 개체 왼쪽 위 = 0,0)")
    for s in slots:
        a(
            f"    slotX({s.index}) = {num(s.x)}: slotY({s.index}) = {num(s.y)}: "
            f"slotW({s.index}) = {num(s.w)}: slotH({s.index}) = {num(s.h)}   ' {s.row + 1}행 {s.col + 1}열"
            + (" (칸에 맞춰 축소)" if s.shrunk else "")
        )
    a("")

    # ---- 기준 개체 / 사진 확보 ----
    if lay.source == SOURCE_SELECTED:
        _emit_selected(a, lay, name)
    else:
        _emit_reference(a)
        if lay.source == SOURCE_FILE_DIALOG:
            _emit_file_dialog(a, lay, name)
        else:
            _emit_folder(a, lay)

    a("    baseX = refShp.Left")
    a("    baseY = refShp.Top")
    a("")

    # ---- 배치 ----
    a("    ' 배치")
    if lay.source == SOURCE_SELECTED:
        a("    For i = 1 To cnt")
        a("        Set pic = pics(i)")
    else:
        a("    For i = 1 To SLOT_COUNT")
        a('        If files(i) <> "" Then')
        a("            Set pic = sld.Shapes.AddPicture( _")
        a("                FileName:=files(i), LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _")
        a("                Left:=baseX + slotX(i) * CM, Top:=baseY + slotY(i) * CM, Width:=-1, Height:=-1)")
    ind = "        " if lay.source == SOURCE_SELECTED else "            "
    a(f"{ind}Fit_{name} pic, baseX + slotX(i) * CM, baseY + slotY(i) * CM, slotW(i) * CM, slotH(i) * CM")
    if lay.border:
        a(f"{ind}With pic.Line")
        a(f"{ind}    .Visible = msoTrue")
        a(f"{ind}    .ForeColor.RGB = {rgb}")
        a(f"{ind}    .Weight = {num(lay.border_weight)}")
        a(f"{ind}    .DashStyle = msoLineSolid")
        a(f"{ind}End With")
    a(f"{ind}pic.ZOrder msoBringToFront")
    if lay.source != SOURCE_SELECTED:
        a("            cnt = cnt + 1")
        a("        End If")
    a("    Next i")
    a("")
    a('    MsgBox "사진 " & cnt & "장 배치 완료!", vbInformation')
    a("")
    a("End Sub")
    a("")

    # ---- 보조 함수 ----
    _emit_fit(a, lay, name)
    if lay.source == SOURCE_FILE_DIALOG and lay.sort == SORT_NATURAL:
        _emit_natcmp(a, name)
    return "\r\n".join(L) + "\r\n"


def _usage(lay: Layout, n: int) -> str:
    if lay.source == SOURCE_SELECTED:
        cnt = f"{n}장" if lay.exact_count else f"최대 {n}장"
        return f"기준 개체(표/도형) 1개 + 사진 {cnt}을 함께 선택 후 실행"
    return "기준 개체(표/도형) 1개를 선택 후 실행"


def _emit_reference(a) -> None:
    a("    ' 기준 개체 1개 확인 (표 셀 안에 커서가 있어도 표 전체를 기준으로 사용)")
    a("    Set sel = ActiveWindow.Selection")
    a("    If sel.Type <> ppSelectionShapes And sel.Type <> ppSelectionText Then")
    a('        MsgBox "기준 개체(표/도형)를 1개 선택한 뒤 실행해주세요.", vbExclamation')
    a("        Exit Sub")
    a("    End If")
    a("    If sel.ShapeRange.Count <> 1 Then")
    a('        MsgBox "기준 개체를 1개만 선택해주세요.", vbExclamation')
    a("        Exit Sub")
    a("    End If")
    a("    Set refShp = sel.ShapeRange(1)")
    a("    Set sld = ActiveWindow.View.Slide")
    a("")


def _count_check(a, lay: Layout, what: str, var: str) -> None:
    if lay.exact_count:
        a(f"    If {var} <> SLOT_COUNT Then")
        a(f'        MsgBox "{what} 정확히 " & SLOT_COUNT & "장 필요합니다. (현재 " & {var} & "장)", vbExclamation')
    else:
        a(f"    If {var} < 1 Or {var} > SLOT_COUNT Then")
        a(f'        MsgBox "{what} 1~" & SLOT_COUNT & "장이어야 합니다. (현재 " & {var} & "장)", vbExclamation')
    a("        Exit Sub")
    a("    End If")


def _emit_file_dialog(a, lay: Layout, name: str) -> None:
    exts = ";".join(f"*.{e}" for e in lay.extensions)
    a("    ' 파일 선택 창에서 사진 다중 선택")
    a("    Set fd = Application.FileDialog(msoFileDialogFilePicker)")
    a("    With fd")
    a(f'        .Title = "배치할 사진을 선택하세요 (" & SLOT_COUNT & "칸)"')
    a("        .AllowMultiSelect = True")
    a("        .Filters.Clear")
    a(f'        .Filters.Add "이미지 파일", {vba_str(exts)}')
    a("        If .Show <> -1 Then Exit Sub")
    a("        cnt = .SelectedItems.Count")
    a("    End With")
    a("")
    _count_check(a, lay, "사진은", "cnt")
    a("    For i = 1 To cnt")
    a("        files(i) = fd.SelectedItems(i)")
    a("    Next i")
    a("")
    if lay.sort == SORT_NATURAL:
        a("    ' 파일명 정렬 (숫자 인식: 1, 2, ..., 10)")
        cmp_expr = f"NatCmp_{name}(FileName_{name}(files(i)), FileName_{name}(files(j)))"
    else:
        a("    ' 파일명 정렬 (문자열 순: 1, 10, 2)")
        cmp_expr = f"StrComp(FileName_{name}(files(i)), FileName_{name}(files(j)), vbTextCompare)"
    a("    For i = 1 To cnt - 1")
    a("        For j = i + 1 To cnt")
    a(f"            If {cmp_expr} > 0 Then")
    a("                tmp = files(i): files(i) = files(j): files(j) = tmp")
    a("            End If")
    a("        Next j")
    a("    Next i")
    a("    cnt = 0")
    a("")


def _emit_folder(a, lay: Layout) -> None:
    ext_list = ", ".join(vba_str("." + e) for e in lay.extensions)
    a("    ' 사진 폴더 선택 → 1, 2, 3... 번호 파일 찾기")
    a("    Set fd = Application.FileDialog(msoFileDialogFolderPicker)")
    a("    With fd")
    a('        .Title = "사진 폴더를 선택하세요 (1, 2, 3... 번호 파일)"')
    a("        If .Show <> -1 Then Exit Sub")
    a("        folder = .SelectedItems(1)")
    a("    End With")
    a('    If Right$(folder, 1) <> "\\" Then folder = folder & "\\"')
    a(f"    exts = Array({ext_list})")
    a("")
    a("    For i = 1 To SLOT_COUNT")
    a("        For k = LBound(exts) To UBound(exts)")
    a('            If Dir$(folder & CStr(i) & exts(k)) <> "" Then')
    a("                files(i) = folder & CStr(i) & exts(k)")
    a("                Exit For")
    a("            End If")
    a("        Next k")
    a('        If files(i) = "" Then missing = missing & " " & i')
    a("    Next i")
    a("")
    if lay.exact_count:
        a('    If missing <> "" Then')
        a('        MsgBox "다음 번호의 사진이 폴더에 없습니다:" & missing, vbExclamation')
        a("        Exit Sub")
        a("    End If")
    else:
        a('    If missing <> "" Then')
        a('        If MsgBox("다음 번호의 사진이 없습니다:" & missing & vbCrLf & "해당 칸은 비워두고 진행할까요?", _')
        a("                  vbQuestion + vbYesNo) = vbNo Then Exit Sub")
        a("    End If")
    a("")


def _emit_selected(a, lay: Layout, name: str) -> None:
    a("    ' 선택 확인: 기준 개체 1개 + 사진")
    a("    Set sel = ActiveWindow.Selection")
    a("    If sel.Type <> ppSelectionShapes Then")
    a('        MsgBox "기준 개체 1개와 사진들을 함께 선택한 뒤 실행해주세요.", vbExclamation')
    a("        Exit Sub")
    a("    End If")
    a("    ReDim pics(1 To sel.ShapeRange.Count)")
    a("    For Each shp In sel.ShapeRange")
    a(f"        If IsPicture_{name}(shp) Then")
    a("            cnt = cnt + 1")
    a("            Set pics(cnt) = shp")
    a("        Else")
    a("            refCount = refCount + 1")
    a("            Set refShp = shp")
    a("        End If")
    a("    Next shp")
    a("    If refCount <> 1 Then")
    a('        MsgBox "사진이 아닌 기준 개체(표/도형)는 1개만 선택해야 합니다. (현재 " & refCount & "개)", vbExclamation')
    a("        Exit Sub")
    a("    End If")
    _count_check(a, lay, "사진은", "cnt")
    a("")
    if lay.selected_order == ORDER_POSITION:
        a("    ' 기존 위치 순서로 정렬: 위→아래, 같은 줄(0.5cm 이내)은 왼쪽→오른쪽")
        a("    For i = 1 To cnt - 1")
        a("        For j = i + 1 To cnt")
        a(f"            If PosAfter_{name}(pics(i), pics(j)) Then")
        a("                Set tmpShp = pics(i): Set pics(i) = pics(j): Set pics(j) = tmpShp")
        a("            End If")
        a("        Next j")
        a("    Next i")
    else:
        a("    ' 선택 순서(ShapeRange 순서)를 그대로 사용")
    a("")


def _emit_fit(a, lay: Layout, name: str) -> None:
    fit = lay.aspect == ASPECT_FIT
    a("' 사진을 박스(x, y, w, h; pt)에 맞춤")
    a(f"Private Sub Fit_{name}(ByVal pic As Shape, ByVal x As Double, ByVal y As Double, _")
    a("                        ByVal w As Double, ByVal h As Double)")
    if fit:
        a("    Dim r As Double, nw As Double, nh As Double")
        a("    ' 원본 비율 유지: 박스 안에 들어가게 축소/확대 후 중앙")
        a("    pic.LockAspectRatio = msoFalse")
        a("    r = pic.Width / pic.Height")
        a("    If r > w / h Then")
        a("        nw = w: nh = w / r")
        a("    Else")
        a("        nh = h: nw = h * r")
        a("    End If")
        a("    pic.Width = nw")
        a("    pic.Height = nh")
        a("    pic.Left = x + (w - nw) / 2")
        a("    pic.Top = y + (h - nh) / 2")
        a("    pic.LockAspectRatio = msoTrue")
    else:
        a("    pic.LockAspectRatio = msoFalse")
        a("    pic.Width = w")
        a("    pic.Height = h")
        a("    pic.Left = x")
        a("    pic.Top = y")
    a("End Sub")
    a("")
    if lay.source == SOURCE_FILE_DIALOG:
        a("' 경로에서 파일명만")
        a(f"Private Function FileName_{name}(ByVal path As String) As String")
        a('    FileName_{0} = Mid$(path, InStrRev(path, "\\") + 1)'.format(name))
        a("End Function")
        a("")
    if lay.source == SOURCE_SELECTED:
        a("' 그림 개체 판별 (그림 개체 틀 포함)")
        a(f"Private Function IsPicture_{name}(ByVal shp As Shape) As Boolean")
        a("    If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then")
        a(f"        IsPicture_{name} = True")
        a("    ElseIf shp.Type = msoPlaceholder Then")
        a("        On Error Resume Next")
        a(f"        IsPicture_{name} = (shp.PlaceholderFormat.ContainedType = msoPicture)")
        a("        On Error GoTo 0")
        a("    End If")
        a("End Function")
        a("")
        if lay.selected_order == ORDER_POSITION:
            a("' a가 b보다 뒤(아래쪽, 같은 줄이면 오른쪽)이면 True")
            a(f"Private Function PosAfter_{name}(ByVal a As Shape, ByVal b As Shape) As Boolean")
            a("    Const ROW_TOL As Double = 0.5 * 72 / 2.54")
            a("    If Abs(a.Top - b.Top) > ROW_TOL Then")
            a(f"        PosAfter_{name} = (a.Top > b.Top)")
            a("    Else")
            a(f"        PosAfter_{name} = (a.Left > b.Left)")
            a("    End If")
            a("End Function")
            a("")


def _emit_natcmp(a, name: str) -> None:
    f = f"NatCmp_{name}"
    a("' 숫자 인식 문자열 비교: a<b → -1, a=b → 0, a>b → 1  (예: 2.jpg < 10.jpg)")
    a(f"Private Function {f}(ByVal a As String, ByVal b As String) As Long")
    a("    Dim i As Long, j As Long")
    a("    Dim ca As String, cb As String")
    a("    Dim na As String, nb As String")
    a("    a = LCase$(a): b = LCase$(b)")
    a("    i = 1: j = 1")
    a("    Do While i <= Len(a) And j <= Len(b)")
    a("        ca = Mid$(a, i, 1): cb = Mid$(b, j, 1)")
    a('        If ca Like "#" And cb Like "#" Then')
    a('            na = "": nb = ""')
    a('            Do While i <= Len(a)')
    a('                If Not (Mid$(a, i, 1) Like "#") Then Exit Do')
    a("                na = na & Mid$(a, i, 1): i = i + 1")
    a("            Loop")
    a('            Do While j <= Len(b)')
    a('                If Not (Mid$(b, j, 1) Like "#") Then Exit Do')
    a("                nb = nb & Mid$(b, j, 1): j = j + 1")
    a("            Loop")
    a('            Do While Len(na) > 1 And Left$(na, 1) = "0": na = Mid$(na, 2): Loop')
    a('            Do While Len(nb) > 1 And Left$(nb, 1) = "0": nb = Mid$(nb, 2): Loop')
    a("            If Len(na) <> Len(nb) Then")
    a(f"                {f} = IIf(Len(na) < Len(nb), -1, 1)")
    a("                Exit Function")
    a("            End If")
    a("            If na <> nb Then")
    a(f"                {f} = IIf(na < nb, -1, 1)")
    a("                Exit Function")
    a("            End If")
    a("        Else")
    a("            If ca <> cb Then")
    a(f"                {f} = StrComp(ca, cb, vbTextCompare)")
    a("                Exit Function")
    a("            End If")
    a("            i = i + 1: j = j + 1")
    a("        End If")
    a("    Loop")
    a(f"    {f} = Sgn((Len(a) - i) - (Len(b) - j))")
    a("End Function")
    a("")
