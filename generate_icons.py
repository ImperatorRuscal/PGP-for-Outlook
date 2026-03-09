#!/usr/bin/env python3
"""
Generate PGP for Outlook add-in icons following MS Office Add-in icon guidelines.
https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-icons

Three variants:
  Icon*        - General / Manage Keys  (blue padlock,  closed)
  IconEncrypt* - Encrypt action         (green padlock, closed)
  IconDecrypt* - Decrypt action         (red padlock,   open)
"""

from PIL import Image, ImageDraw
import os

# ── MS-approved colours ───────────────────────────────────────────────────────
DARK_GRAY   = (58,  58,  56)    # #3A3A38  outline / standalone
MEDIUM_GRAY = (121, 119, 116)   # #797774  secondary content
BG_FILL     = (250, 250, 250)   # #FAFAFA  paper background
LIGHT_GRAY  = (200, 198, 196)   # #C8C6C4  envelope fill

#            standalone          outline            fill
BLUE  = ((30,  139, 205), (0,   99,  177), (131, 190, 236))  # general / keys
RED   = ((237,  61,  59), (212,  35,  20), (255, 145, 152))  # decrypt
GREEN = ((24,  171,  80), (48,  144,  72), (161, 221, 170))  # encrypt


def draw_icon(size: int, lock_sa, lock_ol, lock_fill,
              lock_open: bool = False) -> Image.Image:
    """
    Draw the PGP icon at `size` pixels square.

    Design coordinate space: 80 × 80 units.
    Supersampled 4× for smooth anti-aliasing, then downscaled with LANCZOS.

    Layout:
      • Letter/paper sheet – peeking above the envelope (>=32 px)
      • Envelope body with V-flap
      • Padlock overlay (bottom-right quadrant)
        – Closed (green / blue): centred U-shackle, both legs in body
        – Open   (red):          shackle swung left, right leg only in body
    """
    SS = 4
    W = H = size * SS
    img  = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    def p(v):
        """80-unit design coord → supersampled pixel."""
        return round(v * W / 80)

    def lw(v):
        """Line width, minimum 1 real pixel (= SS supersampled pixels)."""
        return max(SS, round(v * W / 80))

    def rgba(c):
        return (*c, 255)

    def rounded_rect(coords, radius, fill, outline, width):
        try:
            draw.rounded_rectangle(coords, radius=radius,
                                   fill=fill, outline=outline, width=width)
        except AttributeError:
            draw.rectangle(coords, fill=fill, outline=outline, width=width)

    # ── 16 px: simplified (two elements only) ────────────────────────────────
    if size <= 16:
        # Envelope
        draw.rectangle([p(2), p(18), p(60), p(70)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY),
                       width=lw(4))
        mx = p(31)
        draw.line([(p(2),  p(18)), (mx, p(36))], fill=rgba(DARK_GRAY), width=lw(4))
        draw.line([(mx, p(36)), (p(60), p(18))], fill=rgba(DARK_GRAY), width=lw(4))

        # Padlock body (bottom-right badge)
        lk_l, lk_t, lk_r, lk_b = p(46), p(44), p(76), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(4))

        slw = lw(5)
        if lock_open:
            # Shackle: arc extending left of body, right leg only
            s_l, s_t, s_r, s_b = p(26), p(22), p(62), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)
        else:
            # Shackle: centred arc, both legs
            s_l, s_t, s_r, s_b = p(50), p(22), p(72), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)

    # ── 32 px and above: full detailed design ────────────────────────────────
    else:
        # Paper / letter sheet (behind envelope, peeking above flap)
        draw.rectangle([p(7), p(6), p(47), p(50)],
                       fill=rgba(BG_FILL), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Illegible text lines on paper
        for ly in [17, 24, 31, 38]:
            draw.line([(p(11), p(ly)), (p(43), p(ly))],
                      fill=rgba(MEDIUM_GRAY), width=lw(1.5))

        # Envelope body
        draw.rectangle([p(2), p(28), p(54), p(64)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Envelope flap (V-shape)
        mx, fy = p(28), p(40)
        draw.line([(p(2),  p(28)), (mx, fy)], fill=rgba(DARK_GRAY), width=lw(1.5))
        draw.line([(mx, fy), (p(54), p(28))], fill=rgba(DARK_GRAY), width=lw(1.5))

        # Padlock body
        lk_l, lk_t, lk_r, lk_b = p(45), p(42), p(78), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(2))

        # Padlock shackle
        slw = lw(3)
        if lock_open:
            # Open: arc extends LEFT of padlock body; right leg only enters body
            # Free left end floats in air – conventional "open padlock" metaphor
            s_l, s_t, s_r, s_b = p(32), p(22), p(64), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            # Right leg into body
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)
        else:
            # Closed: centred arc, both legs enter padlock body
            s_l, s_t, s_r, s_b = p(51), p(26), p(72), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)

        # Keyhole circle on padlock face
        kh_cx = (lk_l + lk_r) // 2
        kh_cy = (lk_t + lk_b) // 2
        kh_r  = p(5)
        draw.ellipse([kh_cx - kh_r, kh_cy - kh_r,
                      kh_cx + kh_r, kh_cy + kh_r],
                     fill=rgba(lock_sa))

    return img.resize((size, size), Image.LANCZOS)


def main():
    out_dir = os.path.join(os.path.dirname(__file__), 'web', 'images')
    os.makedirs(out_dir, exist_ok=True)

    variants = [
        # (file prefix,   colours, open,  sizes)
        ('Icon',        BLUE,  False, [16, 32, 64, 80, 192]),
        ('IconEncrypt', GREEN, False, [16, 32, 80]),
        ('IconDecrypt', RED,   True,  [16, 32, 80]),
    ]

    for prefix, colours, open_lock, sizes in variants:
        sa, ol, fill = colours
        for sz in sizes:
            img  = draw_icon(sz, sa, ol, fill, open_lock)
            path = os.path.join(out_dir, f'{prefix}{sz}.png')
            img.save(path)
            print(f'  saved  {path}')

    print('Done.')


if __name__ == '__main__':
    main()
