# -*- coding: utf-8 -*-
"""
앱 아이콘(app.ico) 생성 스크립트.
NexaERP 대시보드 톤(인디고)에 맞춘 '접시+포크/나이프+돋보기(검색)' 아이콘.

실행: python make_icon.py  →  app.ico 생성
"""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).with_name("app.ico")

# 팔레트 (GUI와 동일 계열)
INDIGO_TOP = (79, 70, 229)     # #4F46E5
INDIGO_BOT = (49, 46, 129)     # #312E81
WHITE = (255, 255, 255)
PLATE = (238, 240, 250)        # 접시(밝은 회백)
PLATE_RING = (199, 210, 254)   # 접시 테두리
CUTLERY = (255, 255, 255)
LENS_GLASS = (255, 255, 255)
LENS_RING = (255, 255, 255)

SS = 4  # 슈퍼샘플링 배율 (안티에일리어싱)


def rounded_rect_gradient(size, radius):
    """인디고 세로 그라데이션 + 둥근 모서리 배경"""
    grad = Image.new("RGB", (1, size), INDIGO_TOP)
    for y in range(size):
        t = y / max(1, size - 1)
        grad.putpixel((0, y), tuple(
            int(INDIGO_TOP[i] + (INDIGO_BOT[i] - INDIGO_TOP[i]) * t) for i in range(3)
        ))
    grad = grad.resize((size, size))

    mask = Image.new("L", (size, size), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)

    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def draw_icon(px):
    """px 크기의 아이콘 한 장 생성"""
    S = px * SS
    img = rounded_rect_gradient(S, radius=int(S * 0.22))
    d = ImageDraw.Draw(img)

    cx, cy = S * 0.46, S * 0.44
    R = S * 0.24  # 접시 반지름

    # 접시
    d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=PLATE)
    d.ellipse(
        [cx - R * 0.7, cy - R * 0.7, cx + R * 0.7, cy + R * 0.7],
        outline=PLATE_RING, width=max(2, int(S * 0.012)),
    )

    # 포크 (접시 왼쪽)
    fx = cx - R * 1.55
    ftop, fbot = cy - R * 0.95, cy + R * 1.05
    tine_w = max(2, int(S * 0.016))
    for dx in (-S * 0.028, 0, S * 0.028):
        d.line([(fx + dx, ftop), (fx + dx, cy - R * 0.15)], fill=CUTLERY, width=tine_w)
    # 포크 목/자루
    d.line([(fx, cy - R * 0.15), (fx, fbot)], fill=CUTLERY, width=max(3, int(S * 0.03)))
    d.ellipse(
        [fx - S * 0.05, cy - R * 0.28, fx + S * 0.05, cy - R * 0.02],
        fill=CUTLERY,
    )

    # 나이프 (접시 오른쪽)
    kx = cx + R * 1.55
    d.line([(kx, cy - R * 0.95), (kx, fbot)], fill=CUTLERY, width=max(3, int(S * 0.03)))
    d.polygon(
        [
            (kx - S * 0.04, cy - R * 0.95),
            (kx + S * 0.02, cy - R * 0.95),
            (kx + S * 0.02, cy - R * 0.1),
            (kx - S * 0.04, cy),
        ],
        fill=CUTLERY,
    )

    # 돋보기(검색) — 우하단, 접시 위에 겹침
    lx, ly = S * 0.66, S * 0.66
    lr = S * 0.15
    ring_w = max(3, int(S * 0.04))
    # 렌즈 배경(인디고 원)으로 대비
    d.ellipse([lx - lr - ring_w, ly - lr - ring_w, lx + lr + ring_w, ly + lr + ring_w],
              fill=INDIGO_BOT)
    d.ellipse([lx - lr, ly - lr, lx + lr, ly + lr], outline=LENS_RING, width=ring_w)
    d.ellipse([lx - lr * 0.55, ly - lr * 0.55, lx + lr * 0.55, ly + lr * 0.55],
              fill=(255, 255, 255, 60))
    # 손잡이
    hx, hy = lx + lr * 0.75, ly + lr * 0.75
    d.line([(hx, hy), (hx + S * 0.14, hy + S * 0.14)], fill=LENS_RING,
           width=max(4, int(S * 0.055)))

    return img.resize((px, px), Image.LANCZOS)


def main():
    sizes = [16, 24, 32, 48, 64, 128, 256]
    imgs = [draw_icon(s) for s in sizes]
    imgs[0].save(OUT, format="ICO", sizes=[(s, s) for s in sizes], append_images=imgs[1:])
    print(f"아이콘 생성 완료: {OUT}  (크기: {sizes})")

    # 미리보기 PNG도 함께 저장
    draw_icon(256).save(OUT.with_name("app_preview.png"))


if __name__ == "__main__":
    main()
