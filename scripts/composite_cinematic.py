"""
ThepExcelMCP — Cinematic Banner Composite
Strategy: GPT Image 2 cinematic background + Pillow text/chip re-composite on top.
This guarantees zero text garbling while keeping the dramatic GPT aesthetic.

Input:  D:/ThepExcelMCP/assets/banner-cinematic-gpt2-bg_20260626_002234_00.png
Output: D:/ThepExcelMCP/assets/banner-thepexcelmcp-cinematic-1200x630.png
"""

import math
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter

# ---------------------------------------------------------------------------
# Brand tokens
# ---------------------------------------------------------------------------
GOLD         = (212, 168, 75)
GOLD_GLOW    = (255, 220, 100)
GOLD_DIM     = (150, 110, 40)
WHITE        = (255, 255, 255)
ONYX         = (10, 10, 10)
EXCEL_GREEN  = (33, 115, 70)
PQ_TEAL      = (49, 182, 168)
AI_ORANGE    = (217, 119, 87)
BG_CHIP      = (8, 10, 22)     # very dark navy for chip bg overlay

FONT_PATH = Path("C:/Users/sirae/.claude/skills/thepexcel-brand-guidelines/fonts/NotoSansThai-VariableFont_wdth_wght.ttf")
LOGO_PATH = Path("C:/Users/sirae/.claude/skills/thepexcel-brand-guidelines/assets/ThepExcel-Logo-Circle-600x600.png")

BG_PATH  = Path("D:/ThepExcelMCP/assets/banner-cinematic-gpt2-bg_20260626_002234_00.png")
OUT_PATH = Path("D:/ThepExcelMCP/assets/banner-thepexcelmcp-cinematic-1200x630.png")

CAPABILITIES = [
    ("Power Query",    PQ_TEAL,              "PQ"),
    ("PivotTable",     GOLD,                 "PT"),
    ("Chart",          AI_ORANGE,            "CH"),
    ("Data Model/DAX", EXCEL_GREEN,          "DM"),
    ("VBA",            (160, 100, 200),      "VB"),
    ("Python =PY()",   (80, 140, 200),       "PY"),
]


def load_font(size: int) -> ImageFont.FreeTypeFont:
    try:
        return ImageFont.truetype(str(FONT_PATH), size)
    except Exception:
        return ImageFont.load_default()


def text_size(draw, text, font):
    bb = draw.textbbox((0, 0), text, font=font)
    return bb[2] - bb[0], bb[3] - bb[1]


def composite(W: int, H: int):
    # --- Load the cinematic GPT background, resize to target dims ---
    bg = Image.open(BG_PATH).convert("RGBA")
    bg = bg.resize((W, H), Image.LANCZOS)

    # --- Dark scrim over the TEXT ZONES only for legibility ---
    # Left panel: wordmark area (bottom ~40% of left half)
    # Right panel: chips area (right half, full height)
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)

    # Left-bottom scrim (where wordmark lives) — lighter to let bg breathe
    od.rectangle(
        [0, int(H * 0.55), int(W * 0.50), H],
        fill=(6, 8, 20, 130)
    )
    # Right-panel scrim (where chips live) — lighter semi-transparent
    od.rectangle(
        [int(W * 0.50), int(H * 0.10), W, H - int(H * 0.05)],
        fill=(6, 8, 20, 110)
    )

    comp = Image.alpha_composite(bg, overlay)
    draw = ImageDraw.Draw(comp)

    scale = W / 1200

    # --- Fonts ---
    font_wordmark   = load_font(int(72 * scale))
    font_tagline    = load_font(int(22 * scale))
    font_opensource = load_font(int(18 * scale))
    font_chip_label = load_font(int(15 * scale))
    font_chip_badge = load_font(int(13 * scale))
    font_ctrl_label = load_font(int(17 * scale))

    # ---------------------------------------------------------------
    # CHIPS — right panel
    # ---------------------------------------------------------------
    right_margin = int(W * 0.025)
    chips_x      = int(W * 0.515)
    available_w  = W - chips_x - right_margin
    gap_col      = int(W * 0.018)
    cols         = 2
    chip_w       = (available_w - gap_col) // cols
    chip_h       = int(H * 0.115)
    gap          = int(H * 0.028)

    n_rows       = math.ceil(len(CAPABILITIES) / cols)
    total_chip_h = n_rows * chip_h + (n_rows - 1) * gap
    chips_y_start = (H - total_chip_h) // 2

    # "AI now controls" label
    ctrl_text = "AI now controls"
    cw, ch = text_size(draw, ctrl_text, font_ctrl_label)
    total_chip_w = cols * chip_w + (cols - 1) * gap_col
    draw.text(
        (chips_x + (total_chip_w - cw) // 2, chips_y_start - ch - int(H * 0.03)),
        ctrl_text, fill=(*GOLD, 180), font=font_ctrl_label
    )

    for idx, (label, color, badge) in enumerate(CAPABILITIES):
        col = idx % cols
        row = idx // cols
        x = chips_x + col * (chip_w + gap_col)
        y = chips_y_start + row * (chip_h + gap)

        # Chip background (semi-transparent dark + coloured border)
        chip_img = Image.new("RGBA", (chip_w, chip_h), (0, 0, 0, 0))
        cd = ImageDraw.Draw(chip_img)
        cd.rounded_rectangle(
            [0, 0, chip_w - 1, chip_h - 1],
            radius=8,
            fill=(*BG_CHIP, 200),
            outline=(*color, 200),
            width=2
        )
        comp.paste(chip_img, (x, y), chip_img)

        # Badge circle with glow
        badge_r = chip_h // 2 - 8
        bx = x + badge_r + 10
        by = y + chip_h // 2

        # Soft glow behind badge
        glow_r = badge_r + 6
        glow_img = Image.new("RGBA", (glow_r * 2 + 4, glow_r * 2 + 4), (0, 0, 0, 0))
        gd = ImageDraw.Draw(glow_img)
        gd.ellipse([2, 2, glow_r * 2 + 2, glow_r * 2 + 2], fill=(*color, 60))
        glow_img = glow_img.filter(ImageFilter.GaussianBlur(5))
        comp.paste(glow_img, (bx - glow_r - 2, by - glow_r - 2), glow_img)

        draw.ellipse(
            [bx - badge_r, by - badge_r, bx + badge_r, by + badge_r],
            fill=color, outline=WHITE, width=1
        )
        bw, bh = text_size(draw, badge, font_chip_badge)
        draw.text(
            (bx - bw // 2, by - bh // 2), badge,
            fill=WHITE if color != GOLD else ONYX,
            font=font_chip_badge
        )

        # Label text (split on "/" if present)
        parts = label.split("/") if "/" in label else [label]
        if len(parts) == 1 and len(label) > 12:
            sp = label.rfind(" ", 0, len(label) // 2 + 3)
            if sp > 0:
                parts = [label[:sp], label[sp + 1:]]

        label_x = bx + badge_r + 10
        for li, part in enumerate(parts):
            lw, lh = text_size(draw, part, font_chip_label)
            ly_off = -lh // 2 + (li - (len(parts) - 1) / 2) * (lh + 2)
            draw.text(
                (label_x, y + chip_h // 2 + ly_off), part,
                fill=WHITE, font=font_chip_label
            )

    # ---------------------------------------------------------------
    # WORDMARK — lower left
    # ---------------------------------------------------------------
    wm_text = "ThepExcelMCP"
    wm_x = int(W * 0.035)

    # Position: anchored to bottom with fixed margin
    # Calculate total text block height first
    dummy_draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    _, wm_h    = text_size(dummy_draw, wm_text, font_wordmark)
    _, tg_h    = text_size(dummy_draw, "Let AI drive real Excel", font_tagline)
    _, os_h    = text_size(dummy_draw, "Open Source", font_opensource)
    block_h    = wm_h + int(H * 0.022) + tg_h + int(H * 0.025) + os_h + int(H * 0.024)

    bottom_margin = int(H * 0.06)
    wm_y = H - bottom_margin - block_h

    # Gold accent bar above wordmark
    bar_y = wm_y - int(H * 0.018)
    draw.rectangle(
        [wm_x, bar_y, wm_x + int(W * 0.42), bar_y + 3],
        fill=GOLD
    )

    # Drop shadow
    for ox, oy in [(-2, -2), (2, -2), (-2, 2), (2, 2), (0, 3)]:
        draw.text((wm_x + ox, wm_y + oy), wm_text, fill=(0, 0, 0, 200),
                  font=font_wordmark)

    # Three-colour wordmark: Thep=white, Excel=gold, MCP=white
    cur_x = wm_x
    for part_text, part_color in [("Thep", WHITE), ("Excel", GOLD_GLOW), ("MCP", WHITE)]:
        draw.text((cur_x, wm_y), part_text, fill=part_color, font=font_wordmark)
        pw, _ = text_size(draw, part_text, font_wordmark)
        cur_x += pw

    # Tagline
    tg_y = wm_y + wm_h + int(H * 0.022)
    draw.text((wm_x, tg_y), "Let AI drive real Excel",
              fill=(*GOLD, 210), font=font_tagline)

    # Open Source badge
    badge_text = "Open Source  •  MCP Server  •  Windows"
    bw_t, bh_t = text_size(draw, badge_text, font_opensource)
    badge_x = wm_x
    badge_y = tg_y + tg_h + int(H * 0.025)
    pad = int(H * 0.012)
    draw.rounded_rectangle(
        [badge_x - pad, badge_y - pad,
         badge_x + bw_t + pad, badge_y + bh_t + pad],
        radius=6,
        fill=(*EXCEL_GREEN, 60),
        outline=(*EXCEL_GREEN, 170),
        width=1
    )
    draw.text((badge_x, badge_y), badge_text, fill=(*WHITE, 200), font=font_opensource)

    # ---------------------------------------------------------------
    # LOGO (top-left)
    # ---------------------------------------------------------------
    if LOGO_PATH.exists():
        logo_size = int(H * 0.13)
        logo = Image.open(LOGO_PATH).convert("RGBA")
        logo = logo.resize((logo_size, logo_size), Image.LANCZOS)
        lx, ly = int(W * 0.032), int(H * 0.055)
        comp.paste(logo, (lx, ly), logo)

    # Thin separator
    sep_x = int(W * 0.50)
    sep_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sld = ImageDraw.Draw(sep_layer)
    sld.line([(sep_x, int(H * 0.08)), (sep_x, int(H * 0.92))], fill=(*GOLD, 35), width=1)
    comp = Image.alpha_composite(comp, sep_layer)

    # --- Save as RGB PNG ---
    comp.convert("RGB").save(OUT_PATH, "PNG")
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    composite(1200, 630)
    print("Done.")
