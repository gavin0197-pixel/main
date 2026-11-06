"""Layout → 실행 가능한 PowerPoint VBA 코드 생성."""

from __future__ import annotations

from . import APP_NAME, __version__
from . import smartcrop
from .model import (
    ASPECT_FILL,
    ASPECT_FIT,
    ASPECT_STRETCH,
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
    a("' 비율   : " + _aspect_label(lay))
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
    path_arg = '""' if lay.source == SOURCE_SELECTED else "files(i)"
    a(f"{ind}Fit_{name} pic, baseX + slotX(i) * CM, baseY + slotY(i) * CM, slotW(i) * CM, slotH(i) * CM, {path_arg}")
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


def _aspect_label(lay: Layout) -> str:
    if lay.aspect == ASPECT_FIT:
        return "원본 비율 유지 (박스 안에 맞춤, 중앙)"
    if lay.aspect == ASPECT_FILL:
        if lay.smart_crop and lay.source != SOURCE_SELECTED:
            return "원본 비율 유지, 잘라서 꽉 채움 (시료 자동 감지 → 시료가 잘리지 않게)"
        return "원본 비율 유지, 잘라서 꽉 채움 (가운데 기준)"
    return "박스에 강제로 맞춤 (원본 비율이 다르면 찌그러짐)"


def _emit_fit(a, lay: Layout, name: str) -> None:
    mode = lay.aspect
    smart = mode == ASPECT_FILL and lay.smart_crop and lay.source != SOURCE_SELECTED
    a("' 사진을 화면상 박스(x, y, w, h; pt)에 맞춤")
    a("' 휴대폰 사진은 PowerPoint 가 EXIF 방향대로 도형을 90/270도 회전시켜 넣으므로,")
    a("' 회전 전 틀의 가로/세로를 바꿔 지정하고 화면상 박스 중심에 맞춘다.")
    a(f"Private Sub Fit_{name}(ByVal pic As Shape, ByVal x As Double, ByVal y As Double, _")
    a("                        ByVal w As Double, ByVal h As Double, ByVal path As String)")
    a("    Dim rot As Double, sideways As Boolean")
    a("    Dim tw As Double, th As Double   ' 회전 전 틀 기준 목표 크기")
    if mode == ASPECT_FIT:
        a("    Dim r As Double")
    if mode == ASPECT_FILL:
        a("    Dim ow As Double, oh As Double   ' 원본 크기")
        a("    Dim hasBox As Boolean")
        a("    Dim bx0 As Double, by0 As Double, bx1 As Double, by1 As Double")
        a("    Dim wx As Double, wy As Double, ww As Double, wh As Double")
        a("    Dim cl As Double, ct As Double, cr As Double, cb As Double")
        a("    Dim cw As Double, ch As Double")
    a("")
    a("    rot = pic.Rotation")
    a("    sideways = (Abs(rot - 90) < 1 Or Abs(rot - 270) < 1)")
    a("    If sideways Then")
    a("        tw = h: th = w")
    a("    Else")
    a("        tw = w: th = h")
    a("    End If")
    a("    pic.LockAspectRatio = msoFalse")
    a("")
    if mode == ASPECT_FIT:
        a("    ' 원본 비율 유지: 박스 안에 들어가게")
        a("    r = pic.Width / pic.Height")
        a("    If r > tw / th Then")
        a("        th = tw / r")
        a("    Else")
        a("        tw = th * r")
        a("    End If")
        a("    pic.Width = tw")
        a("    pic.Height = th")
    elif mode == ASPECT_FILL:
        a("    ' 기존 자르기 초기화 후 원본 크기 읽기 (큰 사진의 크기 제한을 피하려고 10%로 읽음)")
        a("    With pic.PictureFormat")
        a("        .CropLeft = 0: .CropTop = 0: .CropRight = 0: .CropBottom = 0")
        a("    End With")
        a("    pic.ScaleHeight 0.1, msoTrue")
        a("    pic.ScaleWidth 0.1, msoTrue")
        a("    ow = pic.Width * 10")
        a("    oh = pic.Height * 10")
        a("")
        if smart:
            a("    ' 시료 위치 감지 (실패하면 가운데 기준)")
            a(f"    hasBox = ObjBox_{name}(path, ow, oh, bx0, by0, bx1, by1)")
        else:
            a("    hasBox = False   ' 가운데 기준")
        a(f"    CropWindow_{name} ow, oh, hasBox, bx0, by0, bx1, by1, tw / th, wx, wy, ww, wh")
        a("")
        a("    ' 자르기 (사진 밖으로 나가는 쪽은 자르지 않음 → 시료가 너무 크면 그 방향은 칸보다 좁게)")
        a("    cl = wx: If cl < 0 Then cl = 0")
        a("    ct = wy: If ct < 0 Then ct = 0")
        a("    cr = ow - wx - ww: If cr < 0 Then cr = 0")
        a("    cb = oh - wy - wh: If cb < 0 Then cb = 0")
        a("    With pic.PictureFormat")
        a("        .CropLeft = cl: .CropTop = ct: .CropRight = cr: .CropBottom = cb")
        a("    End With")
        a("    cw = ow - cl - cr")
        a("    ch = oh - ct - cb")
        a("    If cw / ch > tw / th Then")
        a("        th = tw * ch / cw")
        a("    Else")
        a("        tw = th * cw / ch")
        a("    End If")
        a("    pic.Width = tw")
        a("    pic.Height = th")
    else:
        a("    pic.Width = tw")
        a("    pic.Height = th")
    a("")
    a("    ' 화면상 박스 중심에 맞춤 (Left/Top 은 회전 전 틀 기준)")
    a("    pic.Left = x + w / 2 - tw / 2")
    a("    pic.Top = y + h / 2 - th / 2")
    if mode != ASPECT_STRETCH:
        a("    pic.LockAspectRatio = msoTrue")
    a("End Sub")
    a("")
    if mode == ASPECT_FILL:
        _emit_crop_window(a, name)
        if smart:
            _emit_objbox(a, name)
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


def _emit_crop_window(a, name: str) -> None:
    """smartcrop.crop_window / _place 와 동일한 로직."""
    f, pl = f"CropWindow_{name}", f"Place_{name}"
    a("' 잘라낼 창 계산 (사진 틀 좌표, pt). 시료 박스가 있으면 시료가 잘리지 않게, 없으면 가운데 기준")
    a(f"Private Sub {f}(ByVal ow As Double, ByVal oh As Double, ByVal hasBox As Boolean, _")
    a("        ByVal bx0 As Double, ByVal by0 As Double, ByVal bx1 As Double, ByVal by1 As Double, _")
    a("        ByVal ratio As Double, ByRef wx As Double, ByRef wy As Double, ByRef ww As Double, ByRef wh As Double)")
    a(f"    Const PAD As Double = {smartcrop.PAD}   ' 시료 바깥 여유 (사진 크기 비율)")
    a("    Dim cx As Double, cy As Double")
    a("    Dim px0 As Double, py0 As Double, px1 As Double, py1 As Double")
    a("    ' 사진 안에서 가능한 가장 큰 창")
    a("    If ow / oh > ratio Then")
    a("        ww = oh * ratio: wh = oh")
    a("    Else")
    a("        ww = ow: wh = ow / ratio")
    a("    End If")
    a("    If Not hasBox Then")
    a("        cx = ow / 2: cy = oh / 2")
    a("    Else")
    a("        px0 = bx0 - ow * PAD: If px0 < 0 Then px0 = 0")
    a("        py0 = by0 - oh * PAD: If py0 < 0 Then py0 = 0")
    a("        px1 = bx1 + ow * PAD: If px1 > ow Then px1 = ow")
    a("        py1 = by1 + oh * PAD: If py1 > oh Then py1 = oh")
    a("        If px1 - px0 <= ww And py1 - py0 <= wh Then")
    a("            ' 시료 + 여유가 들어감")
    a("            cx = (px0 + px1) / 2: cy = (py0 + py1) / 2")
    a("        ElseIf bx1 - bx0 <= ww And by1 - by0 <= wh Then")
    a("            ' 시료는 들어감 (여유는 일부)")
    a("            cx = (bx0 + bx1) / 2: cy = (by0 + by1) / 2")
    a("        Else")
    a("            ' 시료가 창보다 큼 → 시료 전체를 담는 창")
    a("            If (px1 - px0) / (py1 - py0) > ratio Then")
    a("                ww = px1 - px0: wh = ww / ratio")
    a("            Else")
    a("                wh = py1 - py0: ww = wh * ratio")
    a("            End If")
    a("            cx = (px0 + px1) / 2: cy = (py0 + py1) / 2")
    a("        End If")
    a("    End If")
    a(f"    wx = {pl}(cx, ww, ow)")
    a(f"    wy = {pl}(cy, wh, oh)")
    a("End Sub")
    a("")
    a(f"Private Function {pl}(ByVal center As Double, ByVal size As Double, ByVal total As Double) As Double")
    a("    If size >= total Then")
    a(f"        {pl} = (total - size) / 2")
    a("    Else")
    a(f"        {pl} = center - size / 2")
    a(f"        If {pl} < 0 Then {pl} = 0")
    a(f"        If {pl} > total - size Then {pl} = total - size")
    a("    End If")
    a("End Function")
    a("")


def _emit_objbox(a, name: str) -> None:
    """smartcrop.detect_box + map_box 와 동일한 로직 (Windows 기본 WIA 로 사진 픽셀 읽기)."""
    f = f"ObjBox_{name}"
    a("' 사진에서 시료(배경 흰 천과 다른 부분: 어두운 부품, 은색 파이프, 색 있는 부품) 범위를 찾아")
    a("' 회전 전 틀 좌표(pt)로 돌려줌. 사진을 못 읽거나 시료를 못 찾으면 False → 가운데 기준 자르기")
    a(f"Private Function {f}(ByVal path As String, ByVal ow As Double, ByVal oh As Double, _")
    a("        ByRef bx0 As Double, ByRef by0 As Double, ByRef bx1 As Double, ByRef by1 As Double) As Boolean")
    a(f"    Const K_MAX_SIDE As Long = {smartcrop.MAX_SIDE}   ' 분석용 축소 크기(px)")
    a(f"    Const K_DARK As Double = {smartcrop.DARK}   ' 배경 밝기의 이 비율 미만이면 시료")
    a(f"    Const K_SAT As Long = {smartcrop.SAT}        ' 채도가 이 값 초과면 시료")
    a(f"    Const K_MIN_FRAC As Double = {smartcrop.MIN_FRAC}  ' 한 줄에 시료 픽셀이 이 비율 이상이어야 인정")
    a("    Dim img As Object, ip As Object, v As Object")
    a("    Dim iw As Long, ih As Long, n As Long, k As Long, ix As Long, iy As Long")
    a("    Dim c As Long, r As Long, g As Long, b As Long, mx As Long, mn As Long")
    a("    Dim lumA() As Long, satA() As Long, hist(0 To 255) As Long")
    a("    Dim colCnt() As Long, rowCnt() As Long")
    a("    Dim acc As Long, bg As Long, darkThr As Double, cMin As Long, rMin As Long")
    a("    Dim x0 As Long, x1 As Long, y0 As Long, y1 As Long")
    a("    Dim sx0 As Double, sy0 As Double, sx1 As Double, sy1 As Double")
    a("    Dim u0 As Double, v0 As Double, u1 As Double, v1 As Double")
    a("    Dim orient As Long, sameDir As Boolean")
    a("")
    a(f"    {f} = False")
    a('    If path = "" Then Exit Function')
    a("    On Error GoTo Fail")
    a('    Set img = CreateObject("WIA.ImageFile")')
    a("    img.LoadFile path")
    a("    orient = 1")
    a("    On Error Resume Next")
    a('    orient = img.Properties("274").Value   \' EXIF 방향 (6: 시계 90도, 8: 시계 270도)')
    a("    On Error GoTo Fail")
    a("    ' 저장된 픽셀 방향과 PowerPoint 틀 방향이 같은지 (가로/세로)")
    a("    sameDir = ((img.Width >= img.Height) = (ow >= oh))")
    a("")
    a("    ' 분석용으로 축소")
    a('    Set ip = CreateObject("WIA.ImageProcess")')
    a('    ip.Filters.Add ip.FilterInfos("Scale").FilterID')
    a('    ip.Filters(1).Properties("MaximumWidth").Value = K_MAX_SIDE')
    a('    ip.Filters(1).Properties("MaximumHeight").Value = K_MAX_SIDE')
    a('    ip.Filters(1).Properties("PreserveAspectRatio").Value = True')
    a("    Set img = ip.Apply(img)")
    a("    iw = img.Width: ih = img.Height: n = iw * ih")
    a("    If iw < 4 Or ih < 4 Then Exit Function")
    a("    Set v = img.ARGBData")
    a("")
    a("    ' 픽셀별 밝기/채도, 밝기 분포")
    a("    ReDim lumA(1 To n): ReDim satA(1 To n)")
    a("    For k = 1 To n")
    a("        c = v.Item(k) And &HFFFFFF")
    a("        r = c \\ &H10000: g = (c \\ &H100) And &HFF: b = c And &HFF")
    a("        mx = r: If g > mx Then mx = g")
    a("        If b > mx Then mx = b")
    a("        mn = r: If g < mn Then mn = g")
    a("        If b < mn Then mn = b")
    a("        lumA(k) = (r + g + b) \\ 3")
    a("        satA(k) = mx - mn")
    a("        hist(lumA(k)) = hist(lumA(k)) + 1")
    a("    Next k")
    a("")
    a("    ' 배경 밝기 = 밝기 상위 10% 지점 (흰 천)")
    a("    acc = 0: bg = 255")
    a("    For k = 0 To 255")
    a("        acc = acc + hist(k)")
    a("        If acc >= n * 0.9 Then bg = k: Exit For")
    a("    Next k")
    a("    darkThr = bg * K_DARK")
    a("")
    a("    ' 열/행별 시료 픽셀 수")
    a("    ReDim colCnt(0 To iw - 1): ReDim rowCnt(0 To ih - 1)")
    a("    For iy = 0 To ih - 1")
    a("        For ix = 0 To iw - 1")
    a("            k = iy * iw + ix + 1")
    a("            If lumA(k) < darkThr Or satA(k) > K_SAT Then")
    a("                colCnt(ix) = colCnt(ix) + 1")
    a("                rowCnt(iy) = rowCnt(iy) + 1")
    a("            End If")
    a("        Next ix")
    a("    Next iy")
    a("    cMin = Int(ih * K_MIN_FRAC): If cMin < 2 Then cMin = 2")
    a("    rMin = Int(iw * K_MIN_FRAC): If rMin < 2 Then rMin = 2")
    a("    x0 = -1: y0 = -1")
    a("    For ix = 0 To iw - 1")
    a("        If colCnt(ix) >= cMin Then")
    a("            If x0 < 0 Then x0 = ix")
    a("            x1 = ix")
    a("        End If")
    a("    Next ix")
    a("    For iy = 0 To ih - 1")
    a("        If rowCnt(iy) >= rMin Then")
    a("            If y0 < 0 Then y0 = iy")
    a("            y1 = iy")
    a("        End If")
    a("    Next iy")
    a("    If x0 < 0 Or y0 < 0 Then Exit Function")
    a("    sx0 = x0 / iw: sx1 = (x1 + 1) / iw")
    a("    sy0 = y0 / ih: sy1 = (y1 + 1) / ih")
    a("")
    a("    ' 저장 픽셀 좌표 → PowerPoint 틀 좌표")
    a("    If sameDir Then")
    a("        u0 = sx0: v0 = sy0: u1 = sx1: v1 = sy1")
    a("    ElseIf orient = 6 Then")
    a("        u0 = 1 - sy1: v0 = sx0: u1 = 1 - sy0: v1 = sx1")
    a("    ElseIf orient = 8 Then")
    a("        u0 = sy0: v0 = 1 - sx1: u1 = sy1: v1 = 1 - sx0")
    a("    Else")
    a("        Exit Function")
    a("    End If")
    a("    bx0 = u0 * ow: by0 = v0 * oh: bx1 = u1 * ow: by1 = v1 * oh")
    a(f"    {f} = True")
    a("    Exit Function")
    a("Fail:")
    a(f"    {f} = False")
    a("End Function")
    a("")
