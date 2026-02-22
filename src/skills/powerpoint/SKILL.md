---
name: powerpoint
description: General-purpose PowerPoint skill for reading, creating, and modifying presentation slides.
license: MIT
hosts: [powerpoint]
---

# PowerPoint Default Skill

Use this as the default orchestration skill for PowerPoint tasks.

## Operating Loop

1. **Discover** — Call `get_presentation_overview` to understand slide count and text content.
2. **Read** — Use `get_presentation_content` or `get_slide_image` to inspect specific slides.
3. **Execute** — Create, modify, or reorganize slides using the appropriate tool.
4. **Verify** — Confirm what changed (re-read modified slides if needed).
5. **Summarize** — Finish with a concise plain-language summary of what was done.

## High-Level Tool Guidance

| Task                          | Primary Tool               |
| ----------------------------- | -------------------------- |
| Understand presentation       | `get_presentation_overview`|
| Read slide text               | `get_presentation_content` |
| See slide visually            | `get_slide_image`          |
| Read speaker notes            | `get_slide_notes`          |
| Add a text box                | `set_presentation_content` |
| Create a rich formatted slide | `add_slide_from_code`      |
| Replace an existing slide     | `add_slide_from_code` with `replaceSlideIndex` |
| Edit text in a shape          | `update_slide_shape`       |
| Clear all shapes from slide   | `clear_slide`              |
| Copy a slide                  | `duplicate_slide`          |
| Set speaker notes             | `set_slide_notes`          |

## Common Workflows

### Summarize a presentation
1. `get_presentation_overview` → get all slide text
2. Provide a concise summary to the user

### Create a new slide deck
1. `get_presentation_overview` → understand current state
2. `add_slide_from_code` → create each slide with PptxGenJS (title, bullets, tables, images)
3. Confirm total slides created

### Redesign a slide
1. `get_slide_image` → see current visual design
2. `get_presentation_content` → read the text content
3. `add_slide_from_code` with `replaceSlideIndex` → replace with improved design

### Add content to existing slide
1. `get_presentation_content` → read current text
2. `update_slide_shape` → modify existing shape text, OR
3. `set_presentation_content` → add a new text box

## Choosing Between `set_presentation_content` and `add_slide_from_code`

- **`set_presentation_content`**: Quick text box addition. No formatting control. Good for simple annotations.
- **`add_slide_from_code`**: Full PptxGenJS power — text with fonts/colors/sizes, bullet lists, tables, shapes, images. Use this for any slide that needs to look professional.

## Always-On Defaults

- Always discover the presentation structure before any modification.
- Prefer `add_slide_from_code` over `set_presentation_content` for user-facing content.
- Use 0-based slide indices consistently.
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue independent remaining steps.
