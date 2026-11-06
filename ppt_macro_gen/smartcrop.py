"""'잘라서 꽉 채우기' + 시료 자동 감지 알고리즘의 기준(Python) 구현.

생성되는 VBA(ObjBox_* / CropWindow_*)는 이 파일의 로직을 그대로 옮긴 것이며,
tests/test_smartcrop.py 가 이 구현으로 알고리즘을 검증한다. 둘 중 하나를 고치면 다른 쪽도 같이 고칠 것.

좌표: 사진(회전 전 틀) 기준. 정규화 박스는 0~1.
"""

from __future__ import annotations

# ---- 감지 파라미터 (VBA 상수와 동일하게 유지) ----
MAX_SIDE = 200  # 분석용 축소 크기(px, 긴 변)
DARK = 0.75  # 배경 밝기의 75% 미만이면 시료 (검은 부품, 은색 파이프). 천 주름/접힘은 85% 이상
SAT = 60  # 채도(max-min) 60 초과면 시료 (주황 커넥터 등 색 있는 부분)
MIN_FRAC = 0.02  # 한 열/행에서 시료 픽셀이 2% 이상이어야 시료 범위로 인정 (천 주름 잡티 무시)
PAD = 0.04  # 감지 박스 바깥 여유 (사진 크기의 4%, 들어갈 자리가 있을 때만)


def detect_box(argb: list[int], iw: int, ih: int) -> tuple[float, float, float, float] | None:
    """축소 이미지(ARGB, 왼쪽 위부터 행 순서)에서 시료 범위(여유 없음)를 정규화 박스로 반환. 못 찾으면 None."""
    n = iw * ih
    lum = [0] * n
    sat = [0] * n
    hist = [0] * 256
    for k in range(n):
        c = argb[k] & 0xFFFFFF
        r, g, b = (c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF
        mx, mn = max(r, g, b), min(r, g, b)
        lum[k] = (r + g + b) // 3
        sat[k] = mx - mn
        hist[lum[k]] += 1
    # 배경 밝기 = 밝기 상위 10% 지점 (흰 천)
    target, acc, bg = n * 0.9, 0, 255
    for v in range(256):
        acc += hist[v]
        if acc >= target:
            bg = v
            break
    dark_thr = bg * DARK
    col = [0] * iw
    row = [0] * ih
    for py in range(ih):
        for px in range(iw):
            k = py * iw + px
            if lum[k] < dark_thr or sat[k] > SAT:
                col[px] += 1
                row[py] += 1
    cmin = max(2, int(ih * MIN_FRAC))
    rmin = max(2, int(iw * MIN_FRAC))
    xs = [i for i, v in enumerate(col) if v >= cmin]
    ys = [i for i, v in enumerate(row) if v >= rmin]
    if not xs or not ys:
        return None
    return xs[0] / iw, ys[0] / ih, (xs[-1] + 1) / iw, (ys[-1] + 1) / ih


def map_box(box, same_orientation: bool, exif_orientation: int):
    """저장된 픽셀 방향 기준 박스 → PowerPoint 틀(회전 전) 기준 박스.

    PowerPoint 가 EXIF 회전을 도형 회전(Rotation 90/270)으로 처리하면 틀 = 저장 방향 → 그대로.
    픽셀 자체를 돌려 넣었으면(가로/세로 뒤바뀜) EXIF 방향값으로 변환. 알 수 없으면 None.
    """
    x0, y0, x1, y1 = box
    if same_orientation:
        return box
    if exif_orientation == 6:  # 시계 방향 90° 회전해서 보여줌: (x, y) → (1 - y, x)
        return 1 - y1, x0, 1 - y0, x1
    if exif_orientation == 8:  # 시계 방향 270°: (x, y) → (y, 1 - x)
        return y0, 1 - x1, y1, 1 - x0
    return None


def crop_window(ow: float, oh: float, box, ratio: float) -> tuple[float, float, float, float]:
    """사진(ow×oh) 안에서 가로/세로 비 ratio 인 잘라낼 창 (x, y, w, h).

    box = 시료 범위(틀 좌표, 여유 없음). 없으면 가운데 기준으로 자른다.
      1) 사진 안의 가장 큰 창에 시료 + 여유(PAD)가 들어가면 → 그 주변으로
      2) 여유까지는 안 되지만 시료는 들어가면 → 시료 기준으로 (빈 공간 없음)
      3) 시료가 창보다 크면 → 시료 + 여유를 담는 최소 창. 사진 밖으로 나간 부분은 빈 공간
    """
    if ow / oh > ratio:
        ww, wh = oh * ratio, oh
    else:
        ww, wh = ow, ow / ratio
    if box is None:
        cx, cy = ow / 2, oh / 2
    else:
        bx0, by0, bx1, by1 = box
        px0, py0 = max(0.0, bx0 - ow * PAD), max(0.0, by0 - oh * PAD)
        px1, py1 = min(ow, bx1 + ow * PAD), min(oh, by1 + oh * PAD)
        if px1 - px0 <= ww and py1 - py0 <= wh:
            cx, cy = (px0 + px1) / 2, (py0 + py1) / 2
        elif bx1 - bx0 <= ww and by1 - by0 <= wh:
            cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
        else:
            bw, bh = px1 - px0, py1 - py0
            if bw / bh > ratio:
                ww, wh = bw, bw / ratio
            else:
                ww, wh = bh * ratio, bh
            cx, cy = (px0 + px1) / 2, (py0 + py1) / 2
    wx = _place(cx, ww, ow)
    wy = _place(cy, wh, oh)
    return wx, wy, ww, wh


def _place(center: float, size: float, total: float) -> float:
    if size >= total:
        return (total - size) / 2  # 사진보다 큰 축은 가운데 (양쪽 빈 공간)
    return min(max(center - size / 2, 0.0), total - size)


def clamp_crop(ow: float, oh: float, window) -> tuple[float, float, float, float]:
    """창을 실제 자르기 값(왼, 위, 오른, 아래 ≥ 0)으로. 사진 밖으로 나간 부분은 자르지 않는다
    (그 축은 칸보다 좁게 들어가고, 시료는 잘리지 않음). 음수 자르기에 의존하지 않기 위함."""
    wx, wy, ww, wh = window
    return max(0.0, wx), max(0.0, wy), max(0.0, ow - wx - ww), max(0.0, oh - wy - wh)
