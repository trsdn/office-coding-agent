---
name: powerpoint
description: General-purpose PowerPoint skill for reading, creating, and modifying presentation slides.
license: MIT
hosts: [powerpoint]
---

# PowerPoint Default Skill

Use this as the default orchestration skill for PowerPoint tasks.

## Operating Loop

1. **Locate** — Call `get_selected_slides` to know which slide the user is on right now.
2. **Discover** — Call `get_presentation_overview` to understand slide count and text content.
3. **Read** — Use `get_presentation_content`, `get_slide_shapes`, or `get_slide_image` to inspect the current slide.
4. **Plan** — Before creating or modifying, choose the right layout approach (see Layout Variety below).
5. **Execute** — Create, modify, or reorganize slides using the appropriate tool.
6. **Verify** — Use `get_slide_image` to visually inspect the result. Assume there are issues — find them.
7. **Refine** — Fix issues found, then re-verify. Repeat until a full pass reveals no new issues.
8. **Summarize** — Finish with a concise plain-language summary of what was done.

## High-Level Tool Guidance

| Task                          | Primary Tool               |
| ----------------------------- | -------------------------- |
| Understand presentation       | `get_presentation_overview`|
| Read slide text               | `get_presentation_content` |
| See slide visually            | `get_slide_image`          |
| Read speaker notes            | `get_slide_notes`          |
| List shapes with details      | `get_slide_shapes`         |
| List available layouts        | `get_slide_layouts`        |
| Get selected slides           | `get_selected_slides`      |
| Get selected shapes           | `get_selected_shapes`      |
| Add a text box                | `set_presentation_content` |
| Create a rich formatted slide | `add_slide_from_code`      |
| Replace an existing slide     | `add_slide_from_code` with `replaceSlideIndex` |
| Add geometric shape           | `add_geometric_shape`      |
| Add a line/connector          | `add_line`                 |
| Edit text in a shape          | `update_slide_shape` or `set_shape_text` |
| Change shape colors/style     | `update_shape_style`       |
| Move or resize a shape        | `move_resize_shape`        |
| Delete a specific shape       | `delete_shape`             |
| Clear all shapes from slide   | `clear_slide`              |
| Delete a slide                | `delete_slide`             |
| Reorder slides                | `move_slide`               |
| Set slide background color    | `set_slide_background`     |
| Apply a layout to a slide     | `apply_slide_layout`       |
| Copy a slide (text only)      | `duplicate_slide`          |
| Set speaker notes             | `set_slide_notes`          |

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

## Choosing Between `set_presentation_content` and `add_slide_from_code`

- **`set_presentation_content`**: Quick text box addition. No formatting control. Good for simple annotations.
- **`add_slide_from_code`**: Full PptxGenJS power — text with fonts/colors/sizes, bullet lists, tables, shapes, images. Use this for any slide that needs to look professional.

## Common Workflows

### Summarize a presentation
1. `get_presentation_overview` → get all slide text
2. Provide a concise summary to the user

### Create a new slide deck
1. `get_presentation_overview` → understand current state
2. Plan slide mapping: decide layout type for each slide (vary layouts!)
3. `add_slide_from_code` → create each slide with PptxGenJS
4. `get_slide_image` on each slide → verify visual quality
5. Fix issues, re-verify until clean
6. Confirm total slides created

### Redesign a slide
1. `get_selected_slides` → which slide
2. `get_slide_image` → see current visual design
3. `get_presentation_content` → read the text content
4. `get_slide_shapes` → understand shape layout and positions
5. `add_slide_from_code` with `replaceSlideIndex` → first draft
6. `get_slide_image` → inspect result (assume issues exist — find them)
7. Refine and re-verify until polished

### Add content to existing slide
1. `get_presentation_content` → read current text
2. `get_slide_shapes` → understand existing shape layout
3. `update_slide_shape` or `set_shape_text` → modify existing text, OR
4. `set_presentation_content` → add a new text box
5. Verify with `get_slide_image`

### Modify existing presentation (template adaptation)
1. `get_presentation_overview` → full structure
2. `get_slide_image` on each relevant slide → visual analysis
3. `get_slide_layouts` → available layouts in this deck
4. Plan which slides to keep, modify, delete, or add
5. Make structural changes first (delete/add/reorder slides)
6. Then edit content on each slide
7. Verify each modified slide with `get_slide_image`

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
- **Font size issues** — text too small to read in presentation mode (min 16pt for body)
- **Inconsistent styling** — different fonts, colors, or sizes across similar elements
- **Missing content** — did you include everything the user asked for?
- **Leftover placeholder content** — template text that wasn't replaced

## Template Adaptation Pitfalls

When adapting existing slides or creating from templates:

### Content count mismatch
When source content has fewer items than the template expects:
- **Remove excess elements entirely** — don't just clear text from shapes
- Use `delete_shape` to remove unneeded shapes
- Use `get_slide_shapes` to identify what to remove
- Verify visually that layout still works after removal

When source content has more items than space allows:
- Split across multiple slides rather than cramming
- Consider truncating or summarizing to fit

### Text length mismatch
- **Shorter replacements**: Usually safe
- **Longer replacements**: May overflow text boxes or wrap unexpectedly
- Always verify with `get_slide_image` after text changes
- Adjust font size or box dimensions with `move_resize_shape` if needed

## Formatting Rules

### PptxGenJS best practices (for `add_slide_from_code`)

- **Bold all headings and inline labels**: Use `bold: true` for slide titles, section headers, and labels like "Status:", "Note:"
- **Consistent bullet style**: Use `{ bullet: true }` or `{ bullet: { type: "number" } }` — don't use unicode bullets (•, ‣, etc.)
- **Multi-item content**: Create separate array items for each bullet/paragraph — never concatenate into one string
- **Bold label + description**: NEVER merge into one text run (renders as "LabelDescription" with no space). Always use a colon separator in the same run or put the description on a separate indented line.

**❌ WRONG** — all items in one text element:
```js
slide.addText("Step 1: Do the first thing. Step 2: Do the second thing.", { x: 0.5, y: 2, w: 9, h: 4, fontSize: 18 });
```

**❌ WRONG** — bold label merges into description (renders "Machine LearningSystems that…"):
```js
{ text: "Machine Learning", options: { bold: true, bullet: true } },
{ text: "Systems that learn from data", options: {} },
```

**✅ CORRECT** — separate elements with structure:
```js
slide.addText([
  { text: "Step 1: Do the first thing", options: { bullet: true, fontSize: 16 } },
  { text: "Step 2: Do the second thing", options: { bullet: true, fontSize: 16 } },
], { x: 0.5, y: 2, w: 9, h: 4 });
```

**✅ CORRECT** — bold label with description properly separated:
```js
// Option A: Colon separator in same line (two text runs)
{ text: [
  { text: "Machine Learning: ", options: { bold: true } },
  { text: "Systems that learn from data" }
], options: { bullet: true, fontSize: 14 } },

// Option B: Description on indented sub-line
{ text: "Machine Learning", options: { bold: true, bullet: true, fontSize: 14 } },
{ text: "Systems that learn from data", options: { fontSize: 12, indentLevel: 1 } },
```

- **Color values**: Use 6-digit hex without # prefix: `"4472C4"` not `"#4472C4"`
- **Positioning**: All x, y, w, h values are in inches. Standard slide is 10" × 7.5"
- **Safe margins**: Keep content within 0.5" from slide edges (x ≥ 0.5, y ≥ 0.5, x+w ≤ 9.5, y+h ≤ 7.0)
- **Font sizes**: Title 28–36pt, subtitle 18–22pt, body/bullets 14–16pt, card/column content 11–13pt, table cells 11–13pt
- **Leave 0.3" buffer at bottom** — never fill content to exactly y+h = 7.0"

### Content Limits Per Slide (to prevent overflow)
- **Bullet slides**: Max 5–6 bullets at 14–16pt. Each bullet ≤ 10 words.
- **Multi-column cards**: Max 4 columns. With 4 columns, use 11–12pt and max 3–4 short bullets each.
- **Two-column comparison**: Max 3–4 items per column at 13–14pt.
- **Quote slides**: Max 3 lines of quote text.
- **If content exceeds limits**: Reduce font size, shorten text, or split across multiple slides.

## Always-On Defaults

- **Always call `get_selected_slides` first** to know the user's current slide.
- Always discover the presentation structure before any modification.
- Prefer `add_slide_from_code` over `set_presentation_content` for user-facing content.
- Use 0-based slide indices consistently.
- **Always verify changes with `get_slide_image` after mutations.**
- Vary layouts when creating multi-slide decks.
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue with independent remaining steps.
