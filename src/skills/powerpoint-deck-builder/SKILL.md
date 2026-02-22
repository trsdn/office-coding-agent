---
name: powerpoint-deck-builder
description: >
  Specialized skill for creating new slide decks from scratch. Covers layout variety,
  content sizing, verification loops, and multi-slide deck workflows.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Deck Builder Skill

Activate this skill when creating new presentations or adding multiple slides to an existing deck.

## Deck Creation Workflow

1. `get_presentation_overview` → understand current state
2. **Plan slide mapping**: decide layout type for EACH slide before creating any (see Layout Variety below)
3. `add_slide_from_code` → create each slide with PptxGenJS
4. `get_slide_image` on EVERY slide → verify visual quality
5. Fix issues, re-verify until clean
6. Confirm total slides created

## Layout Variety — CRITICAL

⚠️ **Monotonous presentations are the #1 failure mode.** Do NOT default to title + bullet slides for everything.

When building a deck, actively vary layouts across slides:

- **Title slides** — large title, optional subtitle, minimal elements
- **Bullet lists** — for key points, but keep bullets short (≤8 words per line)
- **Two-column layouts** — comparison, pros/cons, before/after
- **Three-column layouts** — team members, feature cards, process steps
- **Image + text** — hero image on one side, text on the other
- **Full-bleed color** — solid background with centered text for section dividers
- **Stat/number callout** — large number with label for KPIs, metrics, highlights
- **Quote slides** — large quote text with attribution
- **Icon + text rows** — icons with labels for feature overviews
- **Table slides** — structured data in clean tables

**Match content type to layout style:**
- Key points → bullet slide
- Team info → multi-column cards
- Testimonials → quote slide
- Metrics → stat callout
- Process → numbered steps or icon row
- Comparison → two-column side-by-side

**Rule:** In any deck of 5+ slides, never use the same layout pattern for more than 2 consecutive slides.

## Iterative Verification Loop — CRITICAL

**Do not declare success until you've completed at least one fix-and-verify cycle.**

1. Create/modify slides
2. `get_slide_image` → visually inspect the result
3. **List issues found** (if none found, look again more critically)
4. Fix issues
5. **Re-verify affected slides** — one fix often creates another problem
6. Repeat until a full pass reveals no new issues

### What to look for during verification:
- **Overlapping elements** — text through shapes, stacked elements
- **Text overflow** — content cut off at edges or box boundaries
- **Cramped layout** — elements too close together (need breathing room)
- **Uneven spacing** — large empty area in one place, cramped in another
- **Insufficient margins** — content too close to slide edges
- **Poor contrast** — light text on light background, or dark on dark
- **Font size issues** — text too small to read in presentation mode (min 14pt for body)
- **Inconsistent styling** — different fonts, colors, or sizes across similar elements
- **Missing content** — did you include everything the user asked for?

## Content Sizing Rules — CRITICAL

Text overflow (content cut off at box edges) is the #1 visual defect. Follow these rules strictly:

### Font Size Guidelines
| Element | Font Size | Notes |
|---------|-----------|-------|
| Slide title | 28–36pt | One line preferred |
| Subtitle | 18–22pt | |
| Body text / bullets | 14–16pt | Never exceed 18pt for multi-line content |
| Card/column content | 11–13pt | When 3+ columns, use smaller fonts |
| Table cells | 11–13pt | |
| Captions / labels | 10–12pt | |

### Space Budget
- **Safe area**: x ≥ 0.5", y ≥ 0.5", right edge ≤ 9.5", bottom edge ≤ 7.0"
- **Title zone**: y = 0.3–0.5", h = 0.8–1.0" (top 1.5" of slide)
- **Content zone**: y = 1.5–1.8" to y+h ≤ 7.0" (remaining ~5.2" of vertical space)
- **Multi-column**: With N columns, each column width ≈ (9.0 - gaps) / N. Use 0.3" gaps between columns.

### Content Limits Per Slide — MANDATORY
⚠️ **These are hard limits. Exceeding them WILL cause text overflow.**

- **Bullet slides**: Maximum 5 bullets at 14–16pt. Each bullet must be ≤ 8 words total.
- **"Label: Description" bullets**: Keep descriptions SHORT — 3–5 words max after the colon.
  - ✅ `"Machine Learning: Learns from data patterns"` (5 words)
  - ❌ `"Machine Learning: Systems that improve through experience and data"` (TOO LONG)
- **Definition + bullets combo**: Maximum 1-line definition + 4 short bullets. Use 14pt max.
- **Column/card layouts**: Maximum 4 columns. With 4 columns, keep text to 2–3 bullets per column at 11–12pt. Each bullet ≤ 4 words.
- **Two-column comparison**: Maximum 3 items per column at 13–14pt. Each item ≤ 6 words total.
- **Quote slides**: Maximum 3 lines of quote text.
- **General rule**: When in doubt, use FEWER words. Presentations need short punchy text, not full sentences.

### Preventing Overflow
1. **Calculate before coding**: Count your content items and estimate total height BEFORE writing PptxGenJS code.
   - Each line of text at 14pt ≈ 0.3" height. At 12pt ≈ 0.25".
   - A bullet with sub-text (bold title + description) ≈ 0.5–0.7" per item.
2. **If content exceeds space**: Reduce font size, shorten text, remove less important items, or split across slides.
3. **Always leave 0.3" buffer** at the bottom — never fill to exactly y+h = 7.0".
