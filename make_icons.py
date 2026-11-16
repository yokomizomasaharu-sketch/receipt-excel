#!/usr/bin/env python3
"""領収書スキャン＆Excel出力アプリのアイコン（icon-192 / icon-512 / maskable）を生成する。"""
from PIL import Image, ImageDraw, ImageFilter

S = 4            # supersampling factor
BASE = 512       # logical size


def rounded(draw, box, radius, fill=None, outline=None, width=0):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def receipt_path(draw, box, fill, outline, width):
    """ジグザグの切り取り線が下にあるレシート形を描く。"""
    x0, y0, x1, y1 = box
    tooth_h = (y1 - y0) * 0.055
    teeth = 7
    step = (x1 - x0) / teeth
    pts = [(x0, y0), (x1, y0)]
    # 右端を下へ、そのあと下辺をジグザグで左へ
    y_bottom = y1 - tooth_h
    pts.append((x1, y_bottom))
    for i in range(teeth):
        xa = x1 - step * (i + 0.5)
        xb = x1 - step * (i + 1)
        pts.append((xa, y1))
        pts.append((xb, y_bottom))
    pts.append((x0, y0))
    draw.polygon(pts, fill=fill)
    # 輪郭線
    draw.line(pts, fill=outline, width=width, joint='curve')
    return None


def build(size, pad_ratio=0.0):
    W = size * S
    img = Image.new('RGBA', (W, W), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景（角丸スクエア）: 上から下へ淡いグラデーション
    bg = Image.new('RGBA', (W, W), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bg)
    top = (255, 255, 255)
    bot = (226, 234, 244)
    for y in range(W):
        t = y / (W - 1)
        c = tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3))
        bd.line([(0, y), (W, y)], fill=c + (255,))
    mask = Image.new('L', (W, W), 0)
    md = ImageDraw.Draw(mask)
    r = int(W * 0.22)
    md.rounded_rectangle([0, 0, W - 1, W - 1], radius=r, fill=255)
    img.paste(bg, (0, 0), mask)

    # maskable 用の余白（セーフゾーン）
    inset = int(W * pad_ratio)
    box_w = W - inset * 2

    # レシート本体（少し影を落とす）
    rw = int(box_w * 0.50)
    rh = int(box_w * 0.62)
    cx = W // 2
    cy = int(W * 0.47) - int(W * 0.01)
    rx0 = cx - rw // 2
    ry0 = cy - rh // 2
    rx1 = rx0 + rw
    ry1 = ry0 + rh

    shadow = Image.new('RGBA', (W, W), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    receipt_path(sd, (rx0 + int(W * 0.012), ry0 + int(W * 0.018), rx1 + int(W * 0.012), ry1 + int(W * 0.018)),
                 (15, 23, 42, 70), (0, 0, 0, 0), 0)
    shadow = shadow.filter(ImageFilter.GaussianBlur(int(W * 0.012)))
    img.alpha_composite(shadow)

    receipt_path(d, (rx0, ry0, rx1, ry1), (255, 255, 255, 255), (2, 132, 199, 255), max(2, int(W * 0.013)))

    # レシートの文字列（グレーの横線）
    line_c = (148, 163, 184, 255)
    lx0 = rx0 + int(rw * 0.16)
    lw = int(rw * 0.68)
    n = 5
    top_gap = int(rh * 0.16)
    gap = int(rh * 0.115)
    thick = max(2, int(W * 0.011))
    for i in range(n):
        y = ry0 + top_gap + i * gap
        w = lw if i < 3 else int(lw * (0.55 if i == 3 else 0.42))
        d.rounded_rectangle([lx0, y, lx0 + w, y + thick], radius=thick // 2, fill=line_c)

    # 金額行（青で強調）
    y = ry0 + top_gap + 4 * gap + int(rh * 0.02)
    d.rounded_rectangle([lx0, y, lx0 + int(lw * 0.30), y + thick], radius=thick // 2, fill=(2, 132, 199, 255))
    d.rounded_rectangle([lx0 + int(lw * 0.62), y, lx0 + lw, y + int(thick * 1.6)], radius=thick // 2,
                        fill=(2, 132, 199, 255))

    # Excel バッジ（緑の角丸＋白い表）
    bsz = int(W * 0.38)
    bx1 = rx1 + int(W * 0.055)
    by1 = ry1 - int(W * 0.005)
    bx0 = bx1 - bsz
    by0 = by1 - bsz
    bshadow = Image.new('RGBA', (W, W), (0, 0, 0, 0))
    bsd = ImageDraw.Draw(bshadow)
    bsd.rounded_rectangle([bx0, by0 + int(W * 0.014), bx1, by1 + int(W * 0.014)],
                          radius=int(bsz * 0.24), fill=(15, 23, 42, 60))
    bshadow = bshadow.filter(ImageFilter.GaussianBlur(int(W * 0.010)))
    img.alpha_composite(bshadow)

    d.rounded_rectangle([bx0, by0, bx1, by1], radius=int(bsz * 0.24), fill=(22, 163, 74, 255),
                        outline=(255, 255, 255, 255), width=max(2, int(W * 0.012)))

    # バッジ内の表（3x3 グリッド）
    gx0 = bx0 + int(bsz * 0.20)
    gx1 = bx1 - int(bsz * 0.20)
    gy0 = by0 + int(bsz * 0.24)
    gy1 = by1 - int(bsz * 0.24)
    gw = gx1 - gx0
    gh = gy1 - gy0
    gwid = max(2, int(W * 0.009))
    for i in range(3):
        x = gx0 + gw * i / 2
        d.line([(x, gy0), (x, gy1)], fill=(255, 255, 255, 255), width=gwid)
    for j in range(3):
        yy = gy0 + gh * j / 2
        d.line([(gx0, yy), (gx1, yy)], fill=(255, 255, 255, 255), width=gwid)

    out = img.resize((size, size), Image.LANCZOS)
    return out


def main():
    build(512).save('icon-512.png')
    build(192).save('icon-192.png')
    build(512, pad_ratio=0.10).save('icon-maskable-512.png')
    print('icons written')


if __name__ == '__main__':
    main()
