"""시료 자동 감지 + 잘라서 채우기 알고리즘 검증 (실제 사진과 비슷한 합성 사진 사용)."""

import random

import pytest

from ppt_macro_gen.smartcrop import clamp_crop, crop_window, detect_box, map_box

Image = pytest.importorskip("PIL.Image")
ImageDraw = pytest.importorskip("PIL.ImageDraw")

RATIO = 6.5 / 4.5  # 가로 칸


def synth_photo(obj_box, pipe=None, seed=0, size=(1500, 2000)):
    """흰 천 배경(결·주름·그림자) + 검은 브래킷 + 주황 부품 + 은색 파이프. obj_box = 시료 실제 범위(px)."""
    rnd = random.Random(seed)
    W, H = size
    img = Image.new("RGB", size, (236, 236, 238))
    d = ImageDraw.Draw(img)
    for _ in range(400):  # 천 결/주름 (밝은 회색)
        x, y = rnd.randrange(W), rnd.randrange(H)
        v = rnd.randrange(200, 230)
        d.line([(x, y), (x + rnd.randrange(-80, 80), y + rnd.randrange(-20, 20))], fill=(v, v, v + 2), width=3)
    for _ in range(6):  # 큰 접힘 그림자
        y = rnd.randrange(H)
        d.rectangle([0, y, W, y + rnd.randrange(10, 60)], fill=(205, 205, 208))
    x0, y0, x1, y1 = obj_box
    d.rectangle([x0 + 20, y0 + 20, x1 + 25, y1 + 25], fill=(185, 185, 188))  # 시료 그림자
    d.rectangle([x0, y0 + (y1 - y0) // 3, x1, y1], fill=(15, 15, 18))  # 판
    d.rectangle([x0 + (x1 - x0) // 3, y0, x0 + (x1 - x0) // 3 + 60, y1], fill=(10, 10, 12))  # 세운 브래킷
    d.rectangle([x0 - 0, y0 + (y1 - y0) // 2, x0 + 80, y0 + (y1 - y0) // 2 + 120], fill=(225, 110, 45))  # 주황
    if pipe:
        d.rectangle(pipe, fill=(160, 160, 165))
        d.rectangle([pipe[0], pipe[1], pipe[0] + 8, pipe[3]], fill=(90, 90, 95))  # 파이프 어두운 가장자리
    return img


def argb_of(img):
    """WIA 처럼 긴 변 200px 로 축소 후 ARGB 리스트."""
    small = img.copy()
    small.thumbnail((200, 200), Image.BILINEAR)
    w, h = small.size
    raw = small.tobytes()
    data = [(255 << 24) | (raw[i] << 16) | (raw[i + 1] << 8) | raw[i + 2] for i in range(0, len(raw), 3)]
    return data, w, h


def contains(window, box):
    wx, wy, ww, wh = window
    bx0, by0, bx1, by1 = box
    eps = 1e-6
    return wx <= bx0 + eps and wy <= by0 + eps and wx + ww >= bx1 - eps and wy + wh >= by1 - eps


# 사용자 사진과 비슷한 두 종류: 1) 평평한 브래킷 (가운데), 2) 세운 측면 (위아래로 김)
FLAT = dict(obj_box=(70, 380, 1450, 1300))
TALL = dict(obj_box=(100, 130, 1460, 1300), pipe=(780, 1150, 900, 1440))
TRUE_TALL = (100, 130, 1460, 1440)


@pytest.mark.parametrize("case,true_box", [(FLAT, FLAT["obj_box"]), (TALL, TRUE_TALL)])
def test_detected_box_covers_sample(case, true_box):
    img = synth_photo(**case)
    W, H = img.size
    box = detect_box(*argb_of(img))
    assert box is not None
    px = (box[0] * W, box[1] * H, box[2] * W, box[3] * H)
    # 시료 전체(은색 파이프 포함)를 거의 정확히 잡고(축소 오차 ±1px≈10px), 배경을 과하게 잡지 않음
    tol = 12
    assert px[0] <= true_box[0] + tol and px[1] <= true_box[1] + tol
    assert px[2] >= true_box[2] - tol and px[3] >= true_box[3] - tol
    assert (px[3] - px[1]) < (true_box[3] - true_box[1]) + 120


def test_center_crop_would_cut_tall_sample_but_smart_crop_does_not():
    img = synth_photo(**TALL)
    W, H = img.size
    plain = crop_window(W, H, None, RATIO)
    assert not contains(plain, TRUE_TALL)  # 그냥 가운데 자르면 위쪽 브래킷이 잘림
    b = detect_box(*argb_of(img))
    smart = crop_window(W, H, (b[0] * W, b[1] * H, b[2] * W, b[3] * H), RATIO)
    assert contains(smart, TRUE_TALL)
    assert abs(smart[2] / smart[3] - RATIO) < 1e-9
    # 실제 적용되는 자르기(사진 밖은 안 자름): 남는 영역이 시료 전체를 포함, 위아래만 조금 잘림
    cl, ct, cr, cb = clamp_crop(W, H, smart)
    assert (cl, cr) == (0, 0)
    assert contains((cl, ct, W - cl - cr, H - ct - cb), TRUE_TALL)
    assert ct > 0 and cb > 0


def test_flat_sample_fills_cell_without_blank():
    img = synth_photo(**FLAT)
    W, H = img.size
    b = detect_box(*argb_of(img))
    win = crop_window(W, H, (b[0] * W, b[1] * H, b[2] * W, b[3] * H), RATIO)
    assert contains(win, FLAT["obj_box"])
    # 사진 안에서만 잘라냄 (빈 공간 없음), 가로 전체 사용
    assert win[0] >= 0 and win[1] >= 0 and win[0] + win[2] <= W + 1e-6 and win[1] + win[3] <= H + 1e-6
    assert win[2] == pytest.approx(W)


def test_window_math():
    assert crop_window(1500, 2000, None, RATIO) == pytest.approx((0, 1000 - 1500 / RATIO / 2, 1500, 1500 / RATIO))
    # 시료가 창보다 크면 사진 밖(음수 자르기 = 빈 공간)으로 넓힘, 양옆 균등
    # (여유 4% = 세로 80px 포함 → 20~1980)
    wx, wy, ww, wh = crop_window(1500, 2000, (100, 100, 1400, 1900), RATIO)
    assert wh == pytest.approx(1960) and ww == pytest.approx(1960 * RATIO)
    assert wx == pytest.approx((1500 - ww) / 2) and wy == pytest.approx(20)
    # 여유까지는 안 들어가지만 시료는 들어가는 경우 → 빈 공간 없이 시료 기준
    wx, wy, ww, wh = crop_window(1500, 2000, (0, 500, 1500, 1520), RATIO)
    assert (ww, wh) == pytest.approx((1500, 1500 / RATIO)) and wx == 0
    assert wy <= 500 and wy + wh >= 1520
    # 가로로 넓은 사진, 세로 칸
    wx, wy, ww, wh = crop_window(2000, 1500, None, 4.5 / 6.5)
    assert wh == 1500 and ww == pytest.approx(1500 * 4.5 / 6.5)


@pytest.mark.parametrize("powerpoint_rotates_shape", [True, False])
def test_phone_photo_stored_sideways(powerpoint_rotates_shape):
    """휴대폰 사진: 픽셀은 가로로 저장 + EXIF 6(시계 90° 돌려 보기). PowerPoint 처리 방식 2가지 모두."""
    display = synth_photo(**TALL)  # 사람이 보는 세로 사진
    stored = display.rotate(90, expand=True)  # 저장된 픽셀 (반시계 90°)
    sb = detect_box(*argb_of(stored))  # VBA 는 저장 픽셀(WIA)을 분석
    DW, DH = display.size
    if powerpoint_rotates_shape:
        # 틀 = 저장 방향(가로), 도형 Rotation=90. 화면상 가로 칸 → 틀 기준 목표 비율은 뒤집힘
        ow, oh = stored.size
        box = map_box(sb, True, 6)
        win = crop_window(ow, oh, (box[0] * ow, box[1] * oh, box[2] * ow, box[3] * oh), 1 / RATIO)
        # 틀 좌표 → 화면 좌표: (x, y) → (oh - y, x)
        wx, wy, ww, wh = win
        shown = (oh - (wy + wh), wx, wh, ww)
    else:
        # 틀 = 화면 방향(세로)으로 픽셀이 돌아가 들어옴 → EXIF 로 변환
        ow, oh = DW, DH
        box = map_box(sb, False, 6)
        shown = crop_window(ow, oh, (box[0] * ow, box[1] * oh, box[2] * ow, box[3] * oh), RATIO)
    assert contains(shown, TRUE_TALL)
    assert shown[2] / shown[3] == pytest.approx(RATIO)


def test_map_box_orientations():
    b = (0.1, 0.2, 0.3, 0.6)
    assert map_box(b, True, 6) == b
    assert map_box(b, False, 6) == pytest.approx((0.4, 0.1, 0.8, 0.3))
    assert map_box(b, False, 8) == pytest.approx((0.2, 0.7, 0.6, 0.9))
    assert map_box(b, False, 1) is None


def test_blank_image_returns_none():
    img = Image.new("RGB", (300, 400), (240, 240, 240))
    assert detect_box(*argb_of(img)) is None
