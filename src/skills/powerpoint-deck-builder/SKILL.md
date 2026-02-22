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

## Iterative Verification Loop — MANDATORY FOR EVERY SLIDE

**You MUST run this loop for EVERY slide you create. No exceptions.**

### After creating each slide:

```
REPEAT {
  1. Call get_slide_image
  2. Check ALL of these (yes/no for each):
     □ Is ANY text cut off at bottom or sides?
     □ Are ANY words breaking mid-word across lines?
     □ Are elements overlapping?
     □ Is there less than 0.3" margin at bottom?
     □ Is ANY bullet longer than 8 words?
     □ Are column bullets longer than 4 words?
  3. If ANY check fails → FIX with add_slide_from_code + replaceSlideIndex
  4. Go back to step 1
} UNTIL all checks pass
```

### Fix actions by issue type:
| Issue | Fix |
|-------|-----|
| Text cut off at bottom | Remove last bullet OR reduce font 2pt OR shorten all text |
| Word breaking mid-word | Replace long word with shorter synonym. E.g., "Medikamentenentwicklung" → "Arzneimittel", "Verkehrsoptimierung" → "Verkehrsplanung", "Betrugserkennung" → "Betrugsprüfung" |
| Too many bullets | Remove least important bullet. Max 4 if intro paragraph exists |
| Text too long | Rewrite shorter. "Diagnose per Bildanalyse" → "Bilddiagnose" |
| Columns too narrow | Reduce column count (4→3) or shorten ALL bullets to 2 words max |
| Bottom text cut off | Move text up OR reduce content OR increase text box height |

**Long compound words** (common in German, Dutch, Finnish) break layout in narrow columns. During verification, if you see a word split across lines, replace it with a shorter word — even if less precise. Visual clarity > terminological precision.

### Minimum verification calls per deck:
- 5-slide deck = minimum 5 `get_slide_image` calls (one per slide)
- Expect 2-3 fix iterations per slide = 10-15 `get_slide_image` calls total
- **If you call fewer than 5 `get_slide_image` for a 5-slide deck, you skipped verification**

### What to look for during verification:
- **Text overflow** — content cut off at edges or box boundaries (THE #1 DEFECT)
- **Word breaking** — long words split mid-word due to narrow columns
- **Overlapping elements** — text through shapes, stacked elements
- **Cramped layout** — elements too close together (need breathing room)
- **Poor contrast** — light text on light background, or dark on dark
- **Inconsistent styling** — different fonts, colors, or sizes across similar elements

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

- **Bullet-only slides**: Maximum 5 bullets at 14–16pt. Each bullet ≤ 8 words.
- **Definition/intro + bullets**: Maximum 4 bullets (NOT 5!). The intro paragraph eats vertical space — compensate by using fewer bullets.
- **"Label: Description" bullets**: 3–5 words max after the colon.
  - ✅ `"Machine Learning: Lernt aus Datenmustern"` (3 words)
  - ❌ `"Machine Learning: Systeme die durch Erfahrung und Daten verbessert werden"` (TOO LONG)
- **Column/card layouts**: Default 3 columns. 4 only for single-word labels. 12–13pt, 2–3 bullets, ≤ 4 words each. **No word in a column bullet should exceed 12 characters** — replace long words with shorter synonyms.
- **Two-column**: Max 3 items per side at 13–14pt, ≤ 6 words each.
- **Quote slides**: Max 3 lines.
- **General rule**: Fewer words > smaller fonts. Presentations need punchy text, not sentences.

### Preventing Overflow
1. **Calculate before coding**: Count your content items and estimate total height BEFORE writing PptxGenJS code.
   - Each line of text at 14pt ≈ 0.3" height. At 12pt ≈ 0.25".
   - A bullet with sub-text (bold title + description) ≈ 0.5–0.7" per item.
2. **If content exceeds space**: Reduce font size, shorten text, remove less important items, or split across slides.
3. **Always leave 0.3" buffer** at the bottom — never fill to exactly y+h = 7.0".
