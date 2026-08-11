#!/usr/bin/env python3
"""Rebuild Score Examples.pptx using Lead Scoring Overview template styling."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

WORKSPACE = Path(__file__).resolve().parents[1]
TEMPLATE = WORKSPACE / "Lead Scoring Overview - August 2026.pptx"
OUTPUT = WORKSPACE / "Score Examples.pptx"

# Palette from Lead Scoring Overview – August 2026 theme
NAVY = RGBColor(0x07, 0x18, 0x2D)
DARK_BLUE = RGBColor(0x0D, 0x27, 0x4D)
BODY_GRAY = RGBColor(0x41, 0x42, 0x44)
TABLE_GRAY = RGBColor(0x58, 0x59, 0x5B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG = RGBColor(0xF2, 0xF4, 0xF6)
ACCENT_BLUE = RGBColor(0x0A, 0x60, 0xFF)
ACCENT_CYAN = RGBColor(0x02, 0xC8, 0xFF)
ACCENT_ORANGE = RGBColor(0xFF, 0x90, 0x00)
ACCENT_PINK = RGBColor(0xFF, 0x00, 0x7F)
ACCENT_GREEN = RGBColor(0x9F, 0xCC, 0x3B)
WARN_FILL = RGBColor(0xFF, 0xF8, 0xEE)

FONT = "CiscoSansTT"
FONT_LIGHT = "CiscoSansTT Light"

ENGAGEMENT_GRADES = [
    ("4", "76 – 100", "Hot — immediate follow-up"),
    ("3", "51 – 75", "Warm-high — prioritise"),
    ("2", "26 – 50", "Warm — nurture and monitor"),
    ("1", "0 – 25", "Cold — low-touch"),
]


def set_run(font, *, size: int, bold: bool = False, color: RGBColor | None = None, light: bool = False):
    font.name = FONT_LIGHT if light else FONT
    font.size = Pt(size)
    font.bold = bold
    if color:
        font.color.rgb = color


def delete_all_slides(prs: Presentation) -> None:
    for slide_id in list(prs.slides._sldIdLst):
        r_id = slide_id.rId
        prs.part.drop_rel(r_id)
        prs.slides._sldIdLst.remove(slide_id)


def layout_by_name(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"Slide layout not found: {name}")


def set_placeholder_text(slide, idx: int, text: str) -> None:
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            shape.text = text
            return


def add_subtitle(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.35), Inches(0.92), Inches(12.5), Inches(0.38))
    p = box.text_frame.paragraphs[0]
    p.text = text
    set_run(p.font, size=12, color=BODY_GRAY, light=True)


def add_content_slide(prs: Presentation, title: str, subtitle: str = ""):
    slide = prs.slides.add_slide(layout_by_name(prs, "Title Only 1"))
    slide.shapes.title.text = title
    for paragraph in slide.shapes.title.text_frame.paragraphs:
        set_run(paragraph.font, size=24, bold=False, color=NAVY)
    if subtitle:
        add_subtitle(slide, subtitle)
    return slide


def style_table_cell(cell, *, header: bool = False, legend: bool = False, alt: bool = False, font_size: int = 10):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if header:
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_GRAY if not legend else WHITE
        text_color = WHITE if not legend else NAVY
        bold = True
    else:
        if legend:
            cell.fill.solid()
            cell.fill.fore_color.rgb = TABLE_GRAY
            text_color = WHITE
        elif alt:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BG
            text_color = BODY_GRAY
        else:
            text_color = BODY_GRAY
        bold = False

    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER if header or legend else PP_ALIGN.LEFT
        set_run(paragraph.font, size=font_size, bold=bold, color=text_color, light=not bold and not legend)


def add_table(
    slide,
    left,
    top,
    width,
    height,
    headers,
    rows,
    col_widths=None,
    font_size=10,
    legend_style: bool = False,
):
    table_shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height)
    table = table_shape.table
    if col_widths:
        for idx, w in enumerate(col_widths):
            table.columns[idx].width = w

    for col, header in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = header
        style_table_cell(cell, header=True, legend=legend_style, font_size=font_size)

    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(value)
            style_table_cell(
                cell,
                legend=legend_style,
                alt=row_idx % 2 == 0 and not legend_style,
                font_size=font_size,
            )
    return table_shape


def add_outcome_banner(slide, left, top, width, text: str, accent: RGBColor):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.62))
    shape.fill.solid()
    shape.fill.fore_color.rgb = accent
    shape.line.fill.background()
    p = shape.text_frame.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    set_run(p.font, size=13, bold=True, color=WHITE)


def add_callout(slide, left, top, width, height, text: str):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = WARN_FILL
    shape.line.color.rgb = ACCENT_ORANGE
    tf = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), height - Inches(0.24)).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    set_run(p.font, size=12, color=BODY_GRAY, light=True)


def add_score_summary(
    slide,
    left,
    top,
    total: float,
    level: str,
    detail: str,
    banner_text: str,
    outcome_color: RGBColor,
):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(3.55), Inches(1.95))
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BG
    card.line.color.rgb = ACCENT_CYAN

    tf = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.12), Inches(3.2), Inches(1.65)).text_frame
    tf.clear()
    lines = [
        (f"Total engagement score: {total:g}", 15, True, NAVY),
        (f"Engagement level: {level}", 14, True, ACCENT_BLUE),
        ("", 6, False, BODY_GRAY),
        (detail, 12, False, BODY_GRAY),
    ]
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.space_after = Pt(4)
        set_run(p.font, size=size, bold=bold, color=color, light=not bold)

    add_outcome_banner(slide, left, top + Inches(2.08), Inches(3.55), banner_text, outcome_color)


def add_formula_box(slide, left, top, width, height):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = ACCENT_CYAN

    tf = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3)).text_frame
    tf.clear()
    lines = [
        ("How each activity is scored", 16, True, NAVY),
        ("", 4, False, BODY_GRAY),
        ("Points = Activity score % × Category weight %", 14, True, ACCENT_BLUE),
        ("", 4, False, BODY_GRAY),
        ("Example: 90% score × 15% weight = 13.5 points", 12, False, BODY_GRAY),
        ("All qualifying activities in the look-back window are summed.", 12, False, BODY_GRAY),
    ]
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.space_after = Pt(4)
        set_run(p.font, size=size, bold=bold, color=color, light=not bold)


def add_timeline_step_table(slide, steps: list[dict]):
    headers = ["Step", "Activity & rule", "Score %", "Weight", "Points", "Running total"]
    rows = [
        [step["step"], step["activity"], step["score_pct"], step["weight"], step["points"], step["running_total"]]
        for step in steps
    ]
    add_table(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(9.05),
        Inches(min(5.7, 0.42 * (len(rows) + 1))),
        headers,
        rows,
        col_widths=[Inches(0.5), Inches(3.85), Inches(0.8), Inches(0.8), Inches(0.8), Inches(1.05)],
    )


def add_body_text(slide, left, top, width, height, lines, *, lead_bold: bool = False):
    tf = slide.shapes.add_textbox(left, top, width, height).text_frame
    tf.word_wrap = True
    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.space_after = Pt(6)
        set_run(
            p.font,
            size=13 if not lead_bold or i else 12,
            bold=lead_bold and i == 0,
            color=NAVY if lead_bold and i == 0 else BODY_GRAY,
            light=not (lead_bold and i == 0),
        )


def build_deck() -> Presentation:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Template deck not found: {TEMPLATE}")

    prs = Presentation(TEMPLATE)
    delete_all_slides(prs)

    # Slide 1 — Title (matches Overview title slide style)
    title_slide = prs.slides.add_slide(layout_by_name(prs, "Title Slide 1, Two Speakers"))
    set_placeholder_text(title_slide, 0, "Lead Scoring Examples")
    set_placeholder_text(title_slide, 13, "Detailed walkthroughs showing how engagement points are calculated")
    set_placeholder_text(title_slide, 16, "Amanda Chenery")
    set_placeholder_text(title_slide, 12, "August 2026")

    # Slide 2 — How scoring works
    slide = add_content_slide(
        prs,
        "How Engagement Scoring Works",
        "Every activity contributes points; the total determines the engagement level",
    )
    add_formula_box(slide, Inches(0.35), Inches(1.35), Inches(5.45), Inches(2.15))
    add_table(
        slide,
        Inches(6.15),
        Inches(1.35),
        Inches(6.5),
        Inches(2.15),
        ["Level", "Score range", "Meaning"],
        ENGAGEMENT_GRADES,
        col_widths=[Inches(0.65), Inches(1.25), Inches(4.1)],
        legend_style=True,
    )
    add_body_text(
        slide,
        Inches(0.35),
        Inches(3.75),
        Inches(12.5),
        Inches(3.0),
        [
            "Key principles used in every example on the following slides:",
            "• Only activities within the recency window (7, 14, or 30 days) contribute points.",
            "• Higher-intent activities carry a larger category weight (e.g. demo = 25%, email click = 5%).",
            "• Multiple activities stack — the running total is the sum of all qualifying points.",
            "• Leads are routed to VDC when engagement is high enough AND profile fit is A, B, or C.",
        ],
        lead_bold=True,
    )

    # Slide 3 — Example 1
    slide = add_content_slide(
        prs,
        "Example 1: First-Touch Content Syndication",
        "Contact enters Eloqua with only syndicated content activity — low score, no VDC routing",
    )
    add_timeline_step_table(
        slide,
        [
            {
                "step": "1",
                "activity": "Form submit from Integrate (non-hand-raiser) — at least 1 time in last 7 days",
                "score_pct": "90%",
                "weight": "10%",
                "points": "9.0",
                "running_total": "9.0",
            },
            {
                "step": "2",
                "activity": "Content syndication activity created — in last 7 days",
                "score_pct": "90%",
                "weight": "15%",
                "points": "13.5",
                "running_total": "22.5",
            },
        ],
    )
    add_score_summary(
        slide,
        Inches(9.65),
        Inches(1.35),
        22.5,
        "1",
        "Profile must also be A/B/C for any VDC routing.",
        "Low scored — no further action",
        ACCENT_PINK,
    )
    add_body_text(
        slide,
        Inches(0.35),
        Inches(3.25),
        Inches(9.0),
        Inches(1.0),
        [
            "Calculation check: 9.0 + 13.5 = 22.5",
            "Engagement level 1 (0–25) = cold. This lead stays in marketing nurture.",
        ],
    )

    # Slide 4 — Example 2
    slide = add_content_slide(
        prs,
        "Example 2: Multi-Touch with Content Syndication Last Touch",
        "Prior digital activity stacks; last-touch syndication pushes the score above the VDC threshold",
    )
    running = 0.0
    rows = []
    for step_num, activity, score_pct, weight, points in [
        ("1", "Email click-through — in last 14 days", "80%", "5%", 4.0),
        ("2", "Website visit — in last 14 days", "80%", "10%", 8.0),
        ("3", "Landing page visit — in last 14 days", "80%", "15%", 12.0),
        ("4", "Video watched — in last 7 days", "100%", "15%", 15.0),
        ("5", "Webinar watched — in last 7 days", "90%", "10%", 9.0),
        ("6", "Content syndication lead — in last 7 days", "90%", "15%", 13.5),
    ]:
        running += points
        rows.append(
            {
                "step": step_num,
                "activity": activity,
                "score_pct": score_pct,
                "weight": weight,
                "points": f"{points:g}",
                "running_total": f"{running:g}",
            }
        )
    add_timeline_step_table(slide, rows)
    add_score_summary(
        slide,
        Inches(9.65),
        Inches(1.35),
        61.5,
        "2",
        "Requires profile grade A, B, or C.",
        "High scored — passed to VDC",
        ACCENT_BLUE,
    )

    # Slide 5 — Event registration
    slide = add_content_slide(
        prs,
        "Example 3: Event Registration — Step by Step",
        "Each registration adds points; score accumulates as the contact registers for more sessions",
    )
    add_table(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(12.55),
        Inches(2.35),
        ["Step", "Date & event", "Scoring rule", "Score %", "Weight", "Points", "Running total"],
        [
            ("1", "8 May 2024 — Registers for Cisco Live (CLUS)", "Registered at least 1 time in last 7 days", "90%", "10%", "9", "9"),
            ("2", "9 May 2024 — Registers for Keynote session", "Registered more than 1 time in last 7 days", "100%", "10%", "10", "10"),
            ("3", "16 May 2024 — Registers for breakout session", "Registered more than 1 time in last 7 days", "100%", "10%", "10", "10"),
        ],
        col_widths=[Inches(0.45), Inches(2.75), Inches(3.15), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.95)],
        font_size=9,
    )
    add_outcome_banner(
        slide,
        Inches(0.35),
        Inches(4.0),
        Inches(12.55),
        "After step 3: Total = 10 points  •  Engagement level 1  •  No further action (registrations alone)",
        ACCENT_ORANGE,
    )
    add_body_text(
        slide,
        Inches(0.35),
        Inches(4.85),
        Inches(12.55),
        Inches(1.8),
        [
            "Logic → Eloqua identifies each registration transaction and matches it to the Event Registration category.",
            "Calculation → Points = score % × 10% weight. Frequency rules (at least 1 vs. more than 1) determine the score %.",
            "Action → Score accumulates but remains below the high-score threshold until attendance or additional activity occurs.",
        ],
    )

    # Slide 6 — Event attendance
    slide = add_content_slide(
        prs,
        "Example 4: Event Attendance — Step by Step",
        "Attendance carries a higher category weight (30%) and can trigger VDC routing",
    )
    add_table(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(12.55),
        Inches(2.35),
        ["Step", "Date & event", "Scoring rule", "Score %", "Weight", "Points", "Running total"],
        [
            ("1", "2 Jun 2024 — Attends Cisco Live (CLUS)", "Attended at least 1 time in last 7 days", "90%", "30%", "27", "27"),
            ("2", "3 Jun 2024 — Attends Keynote session", "Attended more than 1 time in last 7 days", "100%", "30%", "30", "30"),
            ("3", "4 Jun 2024 — Attends breakout session", "Attended more than 1 time in last 7 days", "100%", "30%", "30", "30"),
        ],
        col_widths=[Inches(0.45), Inches(2.75), Inches(3.15), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.95)],
        font_size=9,
    )
    add_outcome_banner(
        slide,
        Inches(0.35),
        Inches(4.0),
        Inches(12.55),
        "After step 3: Total = 30 points  •  Engagement level 2  •  High-score lead if profile is A",
        ACCENT_GREEN,
    )

    # Slide 7 — Decay
    slide = add_content_slide(
        prs,
        "Example 5: Why Older Registrations Still Count (with Decay)",
        "Registrations from the last 30 days are included, but points are lower outside the 7-day window",
    )
    add_callout(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(12.55),
        Inches(0.85),
        "When a contact attends an event, Eloqua also re-scores earlier registrations from the last 30 days "
        "— but at reduced score percentages due to recency decay.",
    )
    add_table(
        slide,
        Inches(0.35),
        Inches(2.35),
        Inches(12.55),
        Inches(3.05),
        ["Type", "Activity", "Scoring rule", "Score %", "Weight", "Points", "Running total"],
        [
            ("Registration", "8 May — CLUS", "At least 1 time, 30 days", "50%", "10%", "5", "5"),
            ("Registration", "9 May — Keynote", "More than 1 time, 30 days", "60%", "10%", "6", "6"),
            ("Registration", "16 May — Breakout", "More than 1 time, 30 days", "60%", "10%", "6", "6"),
            ("Attendance", "2 Jun — CLUS", "At least 1 time, 7 days", "90%", "30%", "27", "33"),
            ("Attendance", "3 Jun — Keynote", "More than 1 time, 7 days", "100%", "30%", "30", "36"),
            ("Attendance", "4 Jun — Breakout", "More than 1 time, 7 days", "100%", "30%", "30", "36"),
        ],
        col_widths=[Inches(1.0), Inches(2.1), Inches(2.75), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.95)],
        font_size=9,
    )
    add_body_text(
        slide,
        Inches(0.35),
        Inches(5.65),
        Inches(12.55),
        Inches(0.7),
        [
            "Older registrations still contribute at reduced 30-day decay rates. Fresh attendance at 7-day "
            "rates pushes the total to 36. Engagement level 3 → passed to VDC if profile is A.",
        ],
    )

    # Slide 8 — Digital touchpoints
    slide = add_content_slide(
        prs,
        "Example 6: Digital Touchpoints Before Event Registration",
        "Email and web activity that preceded registration is also included in the final score",
    )
    add_callout(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(12.55),
        Inches(0.85),
        "If the contact clicked an email to reach Cisco.com before registering, those transactions "
        "within the last 30 days are included alongside registrations and attendance.",
    )
    add_table(
        slide,
        Inches(0.35),
        Inches(2.35),
        Inches(12.55),
        Inches(2.85),
        ["Type", "Activity", "Scoring rule", "Score %", "Weight", "Points", "Running total"],
        [
            ("Digital", "8 May — Email click-through", "At least 1 time, 30 days", "50%", "5%", "2.5", "2.5"),
            ("Digital", "8 May — Landing page visit", "At least 1 time, 30 days", "50%", "15%", "7.5", "10.0"),
            ("Digital", "8 May — Website visit", "At least 1 time, 30 days", "50%", "10%", "5.0", "15.0"),
            ("Registration", "8 May — CLUS", "At least 1 time, 30 days", "50%", "10%", "5.0", "20.0"),
            ("Registration", "9 May — Keynote", "More than 1 time, 30 days", "60%", "10%", "6.0", "21.0"),
            ("Registration", "16 May — Breakout", "More than 1 time, 30 days", "60%", "10%", "6.0", "21.0"),
        ],
        col_widths=[Inches(1.0), Inches(2.1), Inches(2.75), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.95)],
        font_size=9,
    )

    # Slide 9 — Full journey
    slide = add_content_slide(
        prs,
        "Example 7: Full Journey — All Touchpoints Combined",
        "Complete timeline from first email click through event attendance",
    )
    add_table(
        slide,
        Inches(0.35),
        Inches(1.35),
        Inches(12.55),
        Inches(4.35),
        ["Type", "Activity", "Scoring rule", "Score %", "Weight", "Points", "Running total"],
        [
            ("Digital", "8 May — Email click", "At least 1, 30 days", "50%", "5%", "2.5", "2.5"),
            ("Digital", "8 May — Landing page", "At least 1, 30 days", "50%", "15%", "7.5", "10.0"),
            ("Digital", "8 May — Website visit", "At least 1, 30 days", "50%", "10%", "5.0", "15.0"),
            ("Registration", "8 May — CLUS", "At least 1, 30 days", "50%", "10%", "5.0", "20.0"),
            ("Registration", "9 May — Keynote", "More than 1, 30 days", "60%", "10%", "6.0", "21.0"),
            ("Registration", "16 May — Breakout", "More than 1, 30 days", "60%", "10%", "6.0", "21.0"),
            ("Attendance", "2 Jun — CLUS", "At least 1, 7 days", "90%", "30%", "27.0", "48.0"),
            ("Attendance", "3 Jun — Keynote", "More than 1, 7 days", "100%", "30%", "30.0", "51.0"),
            ("Attendance", "4 Jun — Breakout", "More than 1, 7 days", "100%", "30%", "30.0", "51.0"),
        ],
        col_widths=[Inches(0.95), Inches(1.95), Inches(2.55), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.9)],
        font_size=8,
    )
    add_outcome_banner(
        slide,
        Inches(0.35),
        Inches(5.95),
        Inches(12.55),
        "Final total = 51  •  Engagement level 2  •  Passed to VDC if profile is A, B, or C",
        ACCENT_BLUE,
    )

    # Slide 10 — Summary
    slide = add_content_slide(
        prs,
        "Summary: When Does a Lead Route to VDC?",
        "Engagement score and profile grade must both qualify",
    )
    add_table(
        slide,
        Inches(0.55),
        Inches(1.45),
        Inches(5.65),
        Inches(2.05),
        ["Example", "Engagement total", "Level", "Outcome"],
        [
            ["First-touch syndication", "22.5", "1", "No action"],
            ["Multi-touch syndication", "61.5", "2", "Passed to VDC*"],
            ["Event registration only", "10", "1", "No action"],
            ["Event attendance", "30", "2", "High score if profile A*"],
            ["Full journey", "51", "2", "Passed to VDC*"],
        ],
        col_widths=[Inches(2.15), Inches(1.15), Inches(0.65), Inches(1.35)],
        font_size=9,
    )
    routing = slide.shapes.add_textbox(Inches(6.55), Inches(1.45), Inches(6.2), Inches(5.0)).text_frame
    routing.clear()
    routing_lines = [
        ("Routing rules", 18, True, NAVY),
        ("", 6, False, BODY_GRAY),
        ("* VDC routing also requires:", 14, True, ACCENT_BLUE),
        ("• Profile grade A, B, or C", 13, False, BODY_GRAY),
        ("• Valid CCID present", 13, False, BODY_GRAY),
        ("• Not AU Public Sector (unless exempt)", 13, False, BODY_GRAY),
        ("• Passes Program 389 validation", 13, False, BODY_GRAY),
        ("", 6, False, BODY_GRAY),
        ("Remember:", 14, True, NAVY),
        ("Points = Score % × Weight %", 13, False, BODY_GRAY),
        ("All qualifying activities in the look-back window sum to the total.", 13, False, BODY_GRAY),
        ("Recency decay reduces points for older activity (30-day < 14-day < 7-day).", 13, False, BODY_GRAY),
    ]
    for i, (text, size, bold, color) in enumerate(routing_lines):
        p = routing.paragraphs[0] if i == 0 else routing.add_paragraph()
        p.text = text
        p.space_after = Pt(5)
        set_run(p.font, size=size, bold=bold, color=color, light=not bold)

    return prs


def main():
    deck = build_deck()
    deck.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
