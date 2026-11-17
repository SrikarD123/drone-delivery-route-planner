"""
Generates presentation.pptx — run once, then delete this script.
    python3 make_slides.py
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1A, 0x2E, 0x4A)
BLUE    = RGBColor(0x1F, 0x6F, 0xEB)
ACCENT  = RGBColor(0x00, 0xC2, 0xA0)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY   = RGBColor(0xF4, 0xF6, 0xF9)
DGRAY   = RGBColor(0x55, 0x65, 0x7A)
YELLOW  = RGBColor(0xFF, 0xC8, 0x00)

W = Inches(13.33)   # widescreen width
H = Inches(7.5)     # widescreen height

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]   # completely blank layout


# ── Helpers ───────────────────────────────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill, alpha=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    return shape


def add_text(slide, text, x, y, w, h, size=20, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size    = Pt(size)
    run.font.bold    = bold
    run.font.italic  = italic
    run.font.color.rgb = color
    return tb


def section_header(slide, title, subtitle=None):
    """Dark navy full-bleed header bar at top."""
    add_rect(slide, 0, 0, W, Inches(1.4), NAVY)
    add_text(slide, title, Inches(0.45), Inches(0.18), Inches(12), Inches(0.8),
             size=32, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, subtitle, Inches(0.45), Inches(0.9), Inches(12), Inches(0.45),
                 size=16, color=ACCENT)


def bullet_box(slide, items, x, y, w, h, size=18, title=None, title_color=NAVY):
    if title:
        add_text(slide, title, x, y, w, Inches(0.45),
                 size=20, bold=True, color=title_color)
        y += Inches(0.5)
        h -= Inches(0.5)
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        indent, text = ("    ", item[1:]) if item.startswith("  ") else ("• ", item)
        run.text = indent + text
        run.font.size  = Pt(size)
        run.font.color.rgb = DGRAY if item.startswith("  ") else NAVY


def placeholder_box(slide, x, y, w, h, label="[ Screenshot / Recording Placeholder ]"):
    add_rect(slide, x, y, w, h, LGRAY)
    border = slide.shapes.add_shape(1, x, y, w, h)
    border.fill.background()
    border.line.color.rgb = BLUE
    border.line.width = Pt(1.5)
    add_text(slide, label, x, y + h//2 - Inches(0.25), w, Inches(0.5),
             size=15, italic=True, color=BLUE, align=PP_ALIGN.CENTER)


def code_box(slide, code, x, y, w, h):
    add_rect(slide, x, y, w, h, RGBColor(0x1E, 0x1E, 0x2E))
    tb = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.1),
                                  w - Inches(0.3), h - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = False
    for i, line in enumerate(code.strip().splitlines()):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(12)
        run.font.name = "Courier New"
        run.font.color.rgb = ACCENT


# ─────────────────────────────────────────────────────────────────────────────
# Slide 1 — Title
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, NAVY)
add_rect(sl, 0, Inches(2.6), W, Inches(0.06), ACCENT)

add_text(sl, "Delivery Drone", Inches(1), Inches(0.7), Inches(11), Inches(1.2),
         size=60, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl, "Route Planner", Inches(1), Inches(1.8), Inches(11), Inches(1.0),
         size=60, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(sl, "A Two-Layer AI System: Q-Learning Dispatch + A* Pathfinding",
         Inches(1), Inches(2.85), Inches(11), Inches(0.6),
         size=20, color=RGBColor(0xBB, 0xCC, 0xDD), align=PP_ALIGN.CENTER)

add_text(sl, "CSC 520 — Final Project",
         Inches(1), Inches(3.7), Inches(11), Inches(0.5),
         size=16, color=RGBColor(0xBB, 0xCC, 0xDD), align=PP_ALIGN.CENTER)
add_text(sl, "Srikar Desemsetti",
         Inches(1), Inches(4.15), Inches(11), Inches(0.5),
         size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 2 — What is the Delivery Drone Route Planner?
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, WHITE)
section_header(sl, "What is the Delivery Drone Route Planner?",
               "The problem we set out to solve")

add_text(sl, "Autonomous delivery drones must navigate a city grid, fulfil time-sensitive "
             "delivery jobs, avoid restricted airspace and moving hazards, and manage limited "
             "battery — all without human intervention.",
         Inches(0.45), Inches(1.55), Inches(12.4), Inches(0.9),
         size=17, color=DGRAY)

cols = [Inches(0.45), Inches(4.6), Inches(8.75)]
titles = ["The Environment", "The Challenge", "Our Solution"]
items = [
    ["Discrete city grid (10×50 cells)", "Deliveries with time windows",
     "No-fly zones & charging stations", "Moving obstacles", "Stochastic wind map"],
    ["Where to go next?", "When to recharge vs. deliver?", "How to route around obstacles?",
     "How to generalise across cities?", "How to balance speed vs. safety?"],
    ["Layer 1 — A* plans every route", "Layer 2 — Q-Learning selects targets",
     "Greedy & Random baselines for comparison", "Streamlit frontend for live demos",
     "Benchmark suite across 3 difficulty tiers"],
]
box_colors = [BLUE, RGBColor(0xE6, 0x7E, 0x22), ACCENT]
for col, title, its, bc in zip(cols, titles, items, box_colors):
    add_rect(sl, col, Inches(2.55), Inches(3.9), Inches(0.06), bc)
    add_text(sl, title, col, Inches(2.7), Inches(3.7), Inches(0.45),
             size=16, bold=True, color=bc)
    bullet_box(sl, its, col, Inches(3.2), Inches(3.7), Inches(3.6), size=15)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 3 — Methodology (System Architecture + Algorithm Overview)
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, WHITE)
section_header(sl, "Methodology", "Two-layer AI: Q-Learning dispatch + A* pathfinding")

# Layer boxes
for bx, color, label, sub in [
    (Inches(0.5),  NAVY, "Layer 2 — Q-Learning Agent", "High-level dispatch: WHAT to do"),
    (Inches(6.9),  BLUE, "Layer 1 — A* Pathfinder",    "Low-level routing: HOW to get there"),
]:
    add_rect(sl, bx, Inches(1.7), Inches(5.9), Inches(2.0), color)
    add_text(sl, label, bx + Inches(0.2), Inches(1.9), Inches(5.5), Inches(0.55),
             size=20, bold=True, color=WHITE)
    add_text(sl, sub,   bx + Inches(0.2), Inches(2.45), Inches(5.5), Inches(0.4),
             size=14, color=ACCENT)

# Q-Learning internals
bullet_box(sl, ["MDP state: (pos, battery bin, pending deliveries, time bin)",
                "Actions: DELIVER_i / CHARGE / SKIP_i",
                "Q(s,a) ← Q(s,a) + α·[r + γ·max Q(s',a') − Q(s,a)]",
                "ε-greedy: 1.0 → 0.05 over 5 000 episodes; γ=0.95, α=0.1"],
           Inches(0.7), Inches(3.0), Inches(5.5), Inches(2.2), size=13)

# A* internals
bullet_box(sl, ["f(n) = g(n) + h(n), h = Euclidean distance (admissible)",
                "Edge cost = distance × wind_factor ∈ [0.5, 2.0]",
                "Paths exceeding battery budget pruned during search",
                "replan() reruns A* when moving obstacle blocks route"],
           Inches(7.1), Inches(3.0), Inches(5.5), Inches(2.2), size=13)

# Arrow between boxes
add_text(sl, "target →\n← battery used", Inches(6.05), Inches(2.3), Inches(0.8), Inches(0.9),
         size=11, color=DGRAY, align=PP_ALIGN.CENTER)

# Simulator box at bottom
add_rect(sl, Inches(3.8), Inches(5.5), Inches(5.7), Inches(1.6), LGRAY)
add_rect(sl, Inches(3.8), Inches(5.5), Inches(5.7), Inches(0.06), ACCENT)
add_text(sl, "CityGridSimulator  — stochastic environment (3 tiers: easy / medium / hard)",
         Inches(4.0), Inches(5.6), Inches(5.3), Inches(0.45),
         size=14, bold=True, color=NAVY)
add_text(sl, "Reward: −1/step · +100 on-time delivery · −50 crash  |  100 benchmark scenarios",
         Inches(4.0), Inches(6.1), Inches(5.3), Inches(0.7),
         size=12, color=DGRAY)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 4 — Live Demo (placeholder)
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, WHITE)
section_header(sl, "Live Demo — Streamlit Frontend", "app.py · watch the drone think in real time")

placeholder_box(sl, Inches(0.45), Inches(1.55), Inches(8.3), Inches(5.6),
                "[ Screenshot / Recording of Streamlit App ]")

# Feature list on right
add_rect(sl, Inches(9.1), Inches(1.55), Inches(3.8), Inches(5.6), LGRAY)
add_text(sl, "What you'll see", Inches(9.3), Inches(1.65), Inches(3.4), Inches(0.4),
         size=16, bold=True, color=NAVY)
bullet_box(sl, [
    "Colour-coded grid updates each step",
    "Blue ▲ — drone position",
    "Red — no-fly zones (A* avoids)",
    "Green — charging stations",
    "Yellow → green — deliveries",
    "Orange — moving obstacles",
    "",
    "Battery gauge (green→red)",
    "Step log with rewards",
    "Switch policies mid-run",
    "Auto-play at 1–10 steps/sec",
], Inches(9.1), Inches(2.15), Inches(3.8), Inches(4.8), size=13)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 5 — Benchmark Results
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, WHITE)
section_header(sl, "Benchmark Results", "100 scenarios × 3 tiers × 3 policies")

# Bar chart placeholder
placeholder_box(sl, Inches(0.45), Inches(1.55), Inches(6.2), Inches(5.6),
                "[ results.png — bar chart ]")

# Results table (right)
add_rect(sl, Inches(6.9), Inches(1.55), Inches(6.0), Inches(5.6), LGRAY)
add_text(sl, "Delivery Success Rate", Inches(7.1), Inches(1.65), Inches(5.6), Inches(0.45),
         size=16, bold=True, color=NAVY)

col_headers = ["Policy", "Easy", "Medium", "Hard"]
results = [
    ["Q-Learning", "83.5%", "54.8%", "15.6%"],
    ["Greedy",     "82.9%", "67.9%", "66.1%"],
    ["Random",     "20.0%",  "6.4%",  "1.4%"],
]
highlights = {(1,1), (1,2), (1,3), (2,1), (2,2), (2,3)}   # best-per-col in green
best_per_col = {1: 0, 2: 2, 3: 2}   # row with best value per data col

col_xs = [Inches(7.1), Inches(9.1), Inches(10.5), Inches(11.9)]
for ci, h in enumerate(col_headers):
    add_text(sl, h, col_xs[ci], Inches(2.2), Inches(1.35), Inches(0.4),
             size=14, bold=True, color=NAVY)

for ri, row in enumerate(results):
    y = Inches(2.75 + ri * 0.75)
    if ri == 0:
        add_rect(sl, Inches(6.9), y - Inches(0.05), Inches(6.0), Inches(0.65),
                 RGBColor(0xE8, 0xF4, 0xFF))
    for ci, val in enumerate(row):
        is_best = (ci > 0) and best_per_col.get(ci) == ri
        add_text(sl, val, col_xs[ci], y, Inches(1.35), Inches(0.5),
                 size=15, bold=is_best,
                 color=ACCENT if is_best else DGRAY)

add_rect(sl, Inches(6.9), Inches(5.1), Inches(6.0), Inches(1.8), NAVY)
add_text(sl, "Key Takeaway", Inches(7.1), Inches(5.2), Inches(5.6), Inches(0.4),
         size=14, bold=True, color=ACCENT)
bullet_box(sl,
           ["Q-Learning edges Greedy on easy (learned policy works!)",
            "Greedy wins medium/hard — tabular Q needs more training for large grids",
            "Random is a distant last — confirms AI approaches add real value"],
           Inches(6.9), Inches(5.6), Inches(6.0), Inches(1.2), size=13)
# Override text color in navy box to white
for shape in sl.shapes:
    if hasattr(shape, "text_frame"):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.text.startswith("Q-Learning edges") or \
                   run.text.startswith("Greedy wins") or \
                   run.text.startswith("Random is"):
                    run.font.color.rgb = WHITE


# ─────────────────────────────────────────────────────────────────────────────
# Slide 6 — Discussion (AI Concepts from Class)
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, W, H, WHITE)
section_header(sl, "Discussion",
               "What the results tell us — AI concepts applied")

concepts = [
    (BLUE,  "Search & Heuristics",
     ["A* with admissible/consistent Euclidean heuristic",
      "Completeness guaranteed; optimality guaranteed when h is admissible",
      "Extended with real-world cost: wind-penalized edges, battery pruning"]),
    (RGBColor(0xE6,0x7E,0x22), "Markov Decision Processes",
     ["State, action, transition, reward formally defined in mdp.py",
      "Markov property: next state depends only on current state + action",
      "Discount factor γ encodes preference for sooner vs. later rewards"]),
    (ACCENT, "Reinforcement Learning",
     ["Q-learning: model-free, off-policy temporal-difference learning",
      "Bellman update bootstraps from estimated future value",
      "ε-greedy balances exploration vs. exploitation across training"]),
    (RGBColor(0x8E,0x44,0xAD), "Agent Architecture",
     ["Two-layer hierarchy: deliberative (Q-Learning) + reactive (A*)",
      "High-level actions decomposed into low-level movement primitives",
      "Dynamic replanning demonstrates real-time environment adaptation"]),
]

for i, (color, title, pts) in enumerate(concepts):
    col = Inches(0.45 + (i % 2) * 6.45)
    row = Inches(1.6  + (i // 2) * 2.7)
    add_rect(sl, col, row, Inches(6.0), Inches(2.45), LGRAY)
    add_rect(sl, col, row, Inches(6.0), Inches(0.06), color)
    add_text(sl, title, col + Inches(0.15), row + Inches(0.15),
             Inches(5.7), Inches(0.45), size=17, bold=True, color=color)
    bullet_box(sl, pts, col, row + Inches(0.65), Inches(6.0), Inches(1.7), size=13)


# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
out = "presentation.pptx"
prs.save(out)
print(f"Saved → {out}  ({prs.slides.__len__()} slides)")
