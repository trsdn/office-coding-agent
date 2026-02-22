---
name: Word Document Editing
description: Skill for editing and formatting Microsoft Word documents via the Office JS API.
license: MIT
hosts: [word]
---

# Word Document Editing Skill

Use this as the default orchestration skill for Word document tasks.

## Operating Loop

1. **Discover** — Call `get_document_overview` to understand heading hierarchy, paragraph count, and tables.
2. **Read** — Use `get_document_content`, `get_document_section`, `get_selection_text`, or `get_selection` to inspect the current content.
3. **Execute** — Create, modify, or format content using the appropriate tool.
4. **Verify** — Re-read modified content to confirm changes (e.g. `get_document_section` or `get_selection_text`).
5. **Summarize** — Finish with a concise plain-language summary of what was done.

## High-Level Tool Guidance

| Task                            | Primary Tool                  |
| ------------------------------- | ----------------------------- |
| Understand document structure   | `get_document_overview`       |
| Read full document as HTML      | `get_document_content`        |
| Read a specific section         | `get_document_section`        |
| Get selected text (plain)       | `get_selection_text`          |
| Get selected content (OOXML)    | `get_selection`               |
| Replace entire document         | `set_document_content`        |
| Insert HTML at selection        | `insert_content_at_selection` |
| Insert a paragraph              | `insert_paragraph`            |
| Insert a page/section break     | `insert_break`                |
| Find and replace text           | `find_and_replace`            |
| Insert a table                  | `insert_table`                |
| Insert a bulleted/numbered list | `insert_list`                 |
| Insert an image                 | `insert_image`                |
| Apply font styling to selection | `apply_style_to_selection`    |
| Apply named paragraph style     | `apply_paragraph_style`       |
| Set paragraph formatting        | `set_paragraph_format`        |
| Get document metadata           | `get_document_properties`     |
| Get comments                    | `get_comments`                |
| List content controls           | `get_content_controls`        |
| Insert text at bookmark         | `insert_text_at_bookmark`     |

## Choosing Between Tools

- **`set_document_content`** replaces the ENTIRE document body. Only use when starting fresh or regenerating full content.
- **`insert_content_at_selection`** inserts rich HTML at the cursor. Best for adding formatted content at a specific location.
- **`insert_paragraph`** is simpler — it appends or prepends a paragraph to the document body. Use for quick text additions.
- **`apply_style_to_selection`** changes font-level formatting (bold, italic, size, color). Use for inline text styling.
- **`apply_paragraph_style`** applies a named Word style (e.g. "Heading 1") to paragraphs. Use for structural formatting.
- **`set_paragraph_format`** sets paragraph-level properties (alignment, spacing, indent). Use for layout adjustments.
- **`insert_list`** is the easiest way to create bullet or numbered lists via HTML insertion.

## Iterative Refinement Workflow

Never treat a document change as done after a single pass. Always:

1. **Read** — Inspect the current state before making changes.
2. **Modify** — Apply the requested changes.
3. **Verify** — Re-read the modified content to confirm correctness.
4. **Refine** — If something is off, adjust and verify again.

### What to check during refinement:
- **Formatting** — Are styles, fonts, and spacing correct?
- **Completeness** — Did you include all requested content?
- **Consistency** — Does the new content match the document's existing style?
- **Structure** — Are headings, lists, and tables properly formed?

## Common Workflows

### Summarize a document
1. `get_document_overview` → understand structure
2. `get_document_content` → read full text
3. Provide a concise summary

### Add a new section
1. `get_document_overview` → understand current headings
2. `insert_paragraph` with style "Heading 1" → add heading
3. `insert_content_at_selection` → add section body content
4. Verify with `get_document_section`

### Format existing text
1. `get_selection_text` → see what's selected
2. `apply_style_to_selection` or `apply_paragraph_style` → apply formatting
3. `get_selection_text` → confirm result

### Create a table from data
1. `insert_table` with data array → insert structured table
2. `get_document_content` → verify table appears correctly

## Always-On Defaults

- Always discover document structure before any modification.
- Always read content before modifying it.
- Prefer `insert_content_at_selection` for rich formatted content.
- Use `insert_paragraph` for simple text additions.
- Always verify changes after mutations.
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue with independent remaining steps.
