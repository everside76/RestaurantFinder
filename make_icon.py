# -*- coding: utf-8 -*-
"""
앱 아이콘(app.ico) 생성 스크립트.
NexaERP 대시보드 톤(인디고)에 맞춘 '접시+포크/나이프+돋보기(검색)' 아이콘.

각 크기를 개별 렌더링하여 ICO에 담고, 작은 크기에서는 획을 굵게·요소를
단순화해 또렷하게 보이도록 한다.

실행: python make_icon.py  →  app.ico 생성
"""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).with_name("app.ico")

INDIGO_TOP = (99, 91, 245)     # 약간 밝게
INDIGO_BOT = (49, 46, 129)     # #312E81
PLATE = (244, 245, 252)
PLATE_RING = (165, 180, 252)
CUTLERY = (255, 255, 255)

SS = 8  # 슈퍼샘플링 배율


def rounded_rect_gradient(size, radius):
    grad = Image.new("RGB", (1, size), INDIGO_TOP)
    for y in range(size):
        t = y / max(1, size - 1)
        grad.putpixel((0, y), tuple(
            int(INDIGO_TOP[i] + (INDIGO_BOT[i] - INDIGO_TOP[i]) * t) for i in range(3)
        ))
    grad = grad.resize((size, size))

    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, size - 1, size - 1], radius=radius, fill=255
    )
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def draw_icon(px):
    """px 크기의 아이콘 한 장을 SS배 슈퍼샘플링으로 렌더링 후 축소."""
    small = px <= 32  # 작은 크기는 단순·굵게
    S = px * SS
    img = rounded_rect_gradient(S, radius=int(S * 0.22))
    d = ImageDraw.Draw(img)

    # 작은 크기는 접시를 키우고 돋보기를 크게, 포크 살은 굵게
    cx, cy = S * 0.44, (S * 0.46 if small else S * 0.44)
    R = S * (0.30 if small else 0.25)

    # 접시
    d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=PLATE)
    ring_w = max(2, int(S * (0.018 if small else 0.013)))
    d.ellipse(
        [cx - R * 0.66, cy - R * 0.66, cx + R * 0.66, cy + R * 0.66],
        outline=PLATE_RING, width=ring_w,
    )

    # 포크 (왼쪽)
    fx = cx - R * 1.5
    ftop, fbot = cy - R * 0.9, cy + R * 1.0
    tine_w = max(3, int(S * (0.03 if small else 0.02)))
    stem_w = max(4, int(S * (0.055 if small else 0.038)))
    tine_gap = S * (0.045 if small else 0.032)
    for dx in (-tine_gap, 0, tine_gap):
        d.line([(fx + dx, ftop), (fx + dx, cy - R * 0.25)], fill=CUTLERY, width=tine_w)
    d.line([(fx, cy - R * 0.28), (fx, fbot)], fill=CUTLERY, width=stem_w)

    # 나이프 (오른쪽)
    kx = cx + R * 1.5
    d.line([(kx, cy - R * 0.9), (kx, fbot)], fill=CUTLERY, width=stem_w)
    blade_w = S * (0.055 if small else 0.045)
    d.polygon(
        [(kx - blade_w, cy - R * 0.9), (kx + blade_w * 0.4, cy - R * 0.9),
         (kx + blade_w * 0.4, cy - R * 0.05), (kx - blade_w, cy + R * 0.1)],
        fill=CUTLERY,
    )

    # 돋보기(검색) — 우하단
    lx, ly = S * 0.68, S * 0.68
    lr = S * (0.19 if small else 0.155)
    lens_ring = max(4, int(S * (0.06 if small else 0.045)))
    # 인디고 배경 원으로 대비 확보
    pad = lens_ring + int(S * 0.02)
    d.ellipse([lx - lr - pad, ly - lr - pad, lx + lr + pad, ly + lr + pad], fill=INDIGO_BOT)
    d.ellipse([lx - lr, ly - lr, lx + lr, ly + lr], outline=CUTLERY, width=lens_ring)
    # 손잡이
    hx, hy = lx + lr * 0.72, ly + lr * 0.72
    hlen = S * (0.15 if small else 0.13)
    d.line([(hx, hy), (hx + hlen, hy + hlen)], fill=CUTLERY,
           width=max(5, int(S * (0.075 if small else 0.06))))

    return img.resize((px, px), Image.LANCZOS)


def main():
    sizes = [256, 128, 64, 48, 32, 24, 16]  # 큰 것부터
    imgs = [draw_icon(s) for s in sizes]
    # 가장 큰 이미지를 기준으로, 나머지는 네이티브 렌더를 그대로 첨부
    imgs[0].save(
        OUT, format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=imgs[1:],
    )
    print(f"아이콘 생성 완료: {OUT}")

    # 검증: 실제로 모든 크기가 담겼는지 확인
    chk = Image.open(OUT)
    print("ICO 포함 크기:", sorted(chk.info.get("sizes", []), reverse=True))

    draw_icon(256).save(OUT.with_name("app_preview.png"))


if __name__ == "__main__":
    main()
