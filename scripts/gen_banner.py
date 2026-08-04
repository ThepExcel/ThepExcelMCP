"""
ThepExcelMCP — Facebook Announcement Banner Generator
Technique: Pure Pillow compositing — real rendered text + crisp drawn icons
NO AI lettering — every glyph is rasterized by Pillow from the Noto Sans Thai TTF.

Outputs:
  D:/ThepExcelMCP/assets/banner-thepexcelmcp-1200x630.png   (FB feed 1.91:1)
  D:/ThepExcelMCP/assets/banner-thepexcelmcp-1080x1080.png  (1:1 square)
"""

import math
import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageFilter

# ---------------------------------------------------------------------------
# Brand tokens (ThepExcel brand-guidelines canonical)
# ---------------------------------------------------------------------------
BG_DARK      = (8, 10, 18)       # deep navy-black for premium dark bg
BG_MID       = (14, 18, 32)      # slightly lighter layer
ONYX         = (10, 10, 10)      # #0A0A0A
GOLD         = (212, 168, 75)    # #D4A84B  primary accent
GOLD_DIM     = (150, 110, 40)    # dimmed gold for secondary elements
GOLD_GLOW    = (255, 210, 90)    # brighter gold for glow nodes
WHITE        = (255, 255, 255)
WHITE_80     = (255, 255, 255, 200)
CHARCOAL     = (74, 74, 74)
EXCEL_GREEN  = (33, 115, 70)     # #217346
PQ_TEAL      = (49, 182, 168)    # #31B6A8
AI_ORANGE    = (217, 119, 87)    # #D97757  Anthropic orange

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR   = Path(__file__).parent
PROJECT_ROOT = SCRIPT_DIR.parent
ASSETS_OUT   = PROJECT_ROOT / "assets"
ASSETS_OUT.mkdir(exist_ok=True)

FONT_PATH    = Path("C:/Users/sirae/.claude/skills/thepexcel-brand-guidelines/fonts/NotoSansThai-VariableFont_wdth_wght.ttf")
LOGO_PATH    = Path("C:/Users/sirae/.claude/skills/thepexcel-brand-guidelines/assets/ThepExcel-Logo-Circle-600x600.png")

# ---------------------------------------------------------------------------
# Font helpers
# ---------------------------------------------------------------------------
def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """Load Noto Sans Thai at given size. bold flag adds synthetic weight via size bump for now."""
    try:
        return ImageFont.truetype(str(FONT_PATH), size)
    except Exception:
        return ImageFont.load_default()


def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    bb = draw.textbbox((0, 0), text, font=font)
    return bb[2] - bb[0], bb[3] - bb[1]


# ---------------------------------------------------------------------------
# Background: deep dark with subtle radial glow + grid lines
# ---------------------------------------------------------------------------
def make_background(W: int, H: int) -> Image.Image:
    img = Image.new("RGB", (W, H), BG_DARK)
    draw = ImageDraw.Draw(img)

    # Subtle grid lines
    grid_col = (20, 24, 42)
    step = 60
    for x in range(0, W, step):
        draw.line([(x, 0), (x, H)], fill=grid_col, width=1)
    for y in range(0, H, step):
        draw.line([(0, y), (W, y)], fill=grid_col, width=1)

    # Radial glow centred-left (where the AI core sits)
    cx, cy = int(W * 0.38), H // 2
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    for r in range(340, 0, -4):
        alpha = int(55 * (1 - r / 340) ** 1.6)
        gd.ellipse([cx - r, cy - r, cx + r, cy + r],
                   fill=(212, 168, 75, alpha))
    img = Image.alpha_composite(img.convert("RGBA"), glow).convert("RGB")

    # Secondary green tint glow (Excel nod) — faint, top-right quadrant
    glow2 = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd2 = ImageDraw.Draw(glow2)
    cx2, cy2 = int(W * 0.78), int(H * 0.22)
    for r in range(180, 0, -4):
        alpha = int(28 * (1 - r / 180) ** 2)
        gd2.ellipse([cx2 - r, cy2 - r, cx2 + r, cy2 + r],
                    fill=(33, 115, 70, alpha))
    img = Image.alpha_composite(img.convert("RGBA"), glow2).convert("RGB")

    return img


# ---------------------------------------------------------------------------
# AI core / neural node graphic (drawn, not AI-gen)
# ---------------------------------------------------------------------------
def draw_ai_core(draw: ImageDraw.ImageDraw, cx: int, cy: int, radius: int):
    """Concentric rings + pulsing rays radiating from the AI core centre."""
    # Outer dashed rings
    for i, r in enumerate([radius, int(radius * 0.72), int(radius * 0.48)]):
        alpha_mul = [80, 130, 180][i]
        # Draw ring as many small arcs (simulate dashed)
        n_segments = 60
        for seg in range(n_segments):
            if seg % 3 == 2:
                continue  # gap
            angle_start = seg * (360 / n_segments)
            angle_end   = angle_start + (360 / n_segments) * 0.7
            col = (*GOLD, alpha_mul)
            draw.arc([cx - r, cy - r, cx + r, cy + r],
                     start=angle_start, end=angle_end,
                     fill=GOLD_DIM if i == 0 else GOLD,
                     width=2 if i > 0 else 1)

    # Radiating thin lines
    n_rays = 16
    for i in range(n_rays):
        angle = math.radians(i * 360 / n_rays)
        x1 = cx + int(math.cos(angle) * radius * 0.52)
        y1 = cy + int(math.sin(angle) * radius * 0.52)
        x2 = cx + int(math.cos(angle) * radius * 0.96)
        y2 = cy + int(math.sin(angle) * radius * 0.96)
        draw.line([(x1, y1), (x2, y2)], fill=GOLD_DIM, width=1)

    # Central bright node
    node_r = int(radius * 0.14)
    draw.ellipse([cx - node_r, cy - node_r, cx + node_r, cy + node_r],
                 fill=GOLD_GLOW, outline=WHITE, width=2)
    # Inner dot
    inner_r = int(node_r * 0.45)
    draw.ellipse([cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r],
                 fill=WHITE)


# ---------------------------------------------------------------------------
# Capability icon chips (drawn text + small symbol)
# ---------------------------------------------------------------------------
CAPABILITIES = [
    # (label, accent_color, symbol_char)
    ("Power Query",   PQ_TEAL,      "PQ"),
    ("PivotTable",    GOLD,         "PT"),
    ("Chart",         AI_ORANGE,    "CH"),
    ("Data Model/DAX",EXCEL_GREEN,  "DM"),
    ("VBA",           (160, 100, 200), "VB"),  # purple nod
    ("Python =PY()",  (80, 140, 200),  "PY"),  # blue nod
]


def draw_capability_chips(img: Image.Image, draw: ImageDraw.ImageDraw,
                          chips_x: int, chips_y_start: int,
                          chip_w: int, chip_h: int, gap: int, gap_col: int,
                          font_label: ImageFont.FreeTypeFont,
                          font_badge: ImageFont.FreeTypeFont):
    """Draw 6 capability chips in a 2-column grid."""
    cols = 2
    for idx, (label, color, badge) in enumerate(CAPABILITIES):
        col = idx % cols
        row = idx // cols
        x = chips_x + col * (chip_w + gap_col)
        y = chips_y_start + row * (chip_h + gap)

        # Chip background — dark semi-transparent
        chip_bg = Image.new("RGBA", (chip_w, chip_h), (0, 0, 0, 0))
        cd = ImageDraw.Draw(chip_bg)
        # Rounded-rect chip
        cd.rounded_rectangle([0, 0, chip_w - 1, chip_h - 1],
                              radius=8,
                              fill=(*BG_MID, 210),
                              outline=(*color, 180),
                              width=2)
        img.paste(chip_bg, (x, y), chip_bg)
        draw_ref = draw  # draw on main image

        # Badge circle
        badge_r = chip_h // 2 - 8
        bx = x + badge_r + 10
        by = y + chip_h // 2
        draw.ellipse([bx - badge_r, by - badge_r, bx + badge_r, by + badge_r],
                     fill=color, outline=WHITE, width=1)
        # Badge text
        bw, bh = text_size(draw, badge, font_badge)
        draw.text((bx - bw // 2, by - bh // 2), badge,
                  fill=WHITE if color != GOLD else ONYX,
                  font=font_badge)

        # Label text — split to 2 lines if long
        parts = label.split("/") if "/" in label else [label]
        if len(parts) == 1 and len(label) > 12:
            mid = len(label) // 2
            # split at nearest space
            sp = label.rfind(" ", 0, mid + 3)
            if sp > 0:
                parts = [label[:sp], label[sp + 1:]]
        label_x = bx + badge_r + 10
        line_h = chip_h // 2 - 4 if len(parts) > 1 else chip_h // 2
        for li, part in enumerate(parts):
            lw, lh = text_size(draw, part, font_label)
            ly_offset = -lh // 2 + (li - (len(parts) - 1) / 2) * (lh + 2)
            draw.text((label_x, y + chip_h // 2 + ly_offset), part,
                      fill=WHITE, font=font_label)


# ---------------------------------------------------------------------------
# Connector lines from AI core to capability chips (cinematic feel)
# ---------------------------------------------------------------------------
def draw_connector_lines(img: Image.Image, draw: ImageDraw.ImageDraw,
                         core_cx: int, core_cy: int, core_r: int,
                         chips_x: int, chips_y_start: int,
                         chip_w: int, chip_h: int, gap: int, gap_col: int = 14):
    cols = 2
    for idx in range(len(CAPABILITIES)):
        col = idx % cols
        row = idx // cols
        cx2 = chips_x + col * (chip_w + gap_col) + 14  # left edge of chip
        cy2 = chips_y_start + row * (chip_h + gap) + chip_h // 2
        # Direction from core to chip
        dx = cx2 - core_cx
        dy = cy2 - core_cy
        dist = math.hypot(dx, dy)
        # Start point at core edge
        sx = core_cx + int(dx / dist * core_r)
        sy = core_cy + int(dy / dist * core_r)
        # Draw dashed line
        n_dashes = 12
        for di in range(n_dashes):
            t0 = di / n_dashes
            t1 = (di + 0.55) / n_dashes
            x0 = int(sx + (cx2 - sx) * t0)
            y0 = int(sy + (cy2 - sy) * t0)
            x1 = int(sx + (cx2 - sx) * t1)
            y1 = int(sy + (cy2 - sy) * t1)
            alpha = int(180 * (1 - di / n_dashes))
            draw.line([(x0, y0), (x1, y1)],
                      fill=(*GOLD_DIM, alpha) if idx != 0 else (*PQ_TEAL, alpha),
                      width=1)


# ---------------------------------------------------------------------------
# Main banner compositor
# ---------------------------------------------------------------------------
def make_banner(W: int, H: int, out_path: Path):
    img = make_background(W, H)
    draw = ImageDraw.Draw(img)

    # --- Font sizes scaled to canvas ---
    scale = W / 1200

    # Wordmark fonts
    font_wordmark  = load_font(int(72 * scale))
    font_tagline   = load_font(int(22 * scale))
    font_opensource= load_font(int(18 * scale))
    font_chip_label= load_font(int(int(15 * scale)))
    font_chip_badge= load_font(int(13 * scale))

    # --- AI core position (left-centre-ish) ---
    core_cx = int(W * 0.36)
    core_cy = H // 2
    core_r  = int(min(W, H) * 0.21)

    draw_ai_core(draw, core_cx, core_cy, core_r)

    # --- Capability chips area (right side) ---
    # chips must fit within right panel: from chips_x to W-margin
    right_margin = int(W * 0.025)
    chips_x      = int(W * 0.515)
    available_w  = W - chips_x - right_margin
    gap_col      = int(W * 0.018)
    cols         = 2
    chip_w       = (available_w - gap_col) // cols  # fit both cols in available width
    chip_h       = int(H * 0.115)
    gap          = int(H * 0.028)

    total_chip_w = cols * chip_w + (cols - 1) * gap_col
    n_rows       = math.ceil(len(CAPABILITIES) / cols)
    total_chip_h = n_rows * chip_h + (n_rows - 1) * gap
    chips_y_start = (H - total_chip_h) // 2

    draw_connector_lines(img, draw, core_cx, core_cy, core_r,
                         chips_x, chips_y_start, chip_w, chip_h, gap, gap_col)
    draw_capability_chips(img, draw, chips_x, chips_y_start,
                          chip_w, chip_h, gap, gap_col,
                          font_chip_label, font_chip_badge)

    # --- Logo (top-left) ---
    if LOGO_PATH.exists():
        logo_size = int(H * 0.14)
        logo = Image.open(LOGO_PATH).convert("RGBA")
        logo = logo.resize((logo_size, logo_size), Image.LANCZOS)
        lx = int(W * 0.035)
        ly = int(H * 0.06)
        img.paste(logo, (lx, ly), logo)

    # --- Wordmark "ThepExcelMCP" ---
    wm_text = "ThepExcelMCP"
    wm_w, wm_h = text_size(draw, wm_text, font_wordmark)

    # Place wordmark in lower-left quadrant, comfortably above the bottom
    wm_x = int(W * 0.035)
    # For 1.91:1 (630H): 60% = 378 → fine
    # For 1:1 (1080H): 60% = 648 → too low (leaves dead zone)
    # Use a height-fraction that works for both: tie to core_cy + core_r
    wm_y_candidate = core_cy + core_r + int(H * 0.04)
    wm_y_max = int(H * 0.72)  # never let it go below 72% — leaves room for tagline
    wm_y = min(wm_y_candidate, wm_y_max)

    # Gold accent bar above wordmark
    bar_y = wm_y - int(H * 0.018)
    draw.rectangle([wm_x, bar_y, wm_x + min(wm_w, int(W * 0.46)), bar_y + 3],
                   fill=GOLD)

    # Shadow pass for wordmark legibility
    for ox, oy in [(-2, -2), (2, -2), (-2, 2), (2, 2), (0, 3)]:
        draw.text((wm_x + ox, wm_y + oy), wm_text, fill=(0, 0, 0, 200),
                  font=font_wordmark)
    # Main wordmark — "Thep" white, "Excel" gold, "MCP" white
    # Draw character-by-character for colour split
    parts_wm = [("Thep", WHITE), ("Excel", GOLD), ("MCP", WHITE)]
    cur_x = wm_x
    for part_text, part_color in parts_wm:
        draw.text((cur_x, wm_y), part_text, fill=part_color, font=font_wordmark)
        pw, _ = text_size(draw, part_text, font_wordmark)
        cur_x += pw

    # --- Tagline ---
    tagline = "Let AI drive real Excel"
    tg_w, tg_h = text_size(draw, tagline, font_tagline)
    tg_y = wm_y + wm_h + int(H * 0.022)
    draw.text((wm_x, tg_y), tagline, fill=(*GOLD, 210), font=font_tagline)

    # --- "Open Source MCP Server" badge ---
    badge_text = "Open Source  •  MCP Server  •  Windows"
    bw, bh = text_size(draw, badge_text, font_opensource)
    badge_x = wm_x
    badge_y = tg_y + tg_h + int(H * 0.025)
    # pill bg
    pad = int(H * 0.012)
    draw.rounded_rectangle(
        [badge_x - pad, badge_y - pad,
         badge_x + bw + pad, badge_y + bh + pad],
        radius=6,
        fill=(*EXCEL_GREEN, 60),
        outline=(*EXCEL_GREEN, 160),
        width=1
    )
    draw.text((badge_x, badge_y), badge_text,
              fill=(*WHITE, 200), font=font_opensource)

    # --- "AI Controls" label above chips ---
    ctrl_text = "AI now controls"
    ctrl_font = load_font(int(17 * scale))
    cw, ch = text_size(draw, ctrl_text, ctrl_font)
    draw.text((chips_x + (total_chip_w - cw) // 2,
               chips_y_start - ch - int(H * 0.03)),
              ctrl_text, fill=(*GOLD, 160), font=ctrl_font)

    # --- Thin separator line between left and right panels ---
    sep_x = int(W * 0.50)
    draw.line([(sep_x, int(H * 0.08)), (sep_x, int(H * 0.92))],
              fill=(*GOLD, 40), width=1)

    img.save(out_path, "PNG", optimize=False)
    print(f"Saved: {out_path}")


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    make_banner(1200, 630,  ASSETS_OUT / "banner-thepexcelmcp-1200x630.png")
    make_banner(1080, 1080, ASSETS_OUT / "banner-thepexcelmcp-1080x1080.png")
    print("Done.")
