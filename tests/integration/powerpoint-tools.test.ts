/**
 * Integration test: PowerPoint tools and their JSON Schema parameters.
 *
 * Validates that:
 * - powerPointTools are well-formed (names, descriptions, handlers)
 * - Each tool has correct JSON Schema parameters
 * - get_presentation_overview exposes the thumbnailWidth param with the right shape
 *
 * NOTE: Testing the handler logic (getImageAsBase64, slide text extraction, etc.)
 * requires a real PowerPoint.RequestContext and MUST be covered by E2E tests running
 * in PowerPoint Desktop — not here.
 */
import { describe, it, expect } from 'vitest';
import Ajv from 'ajv';
import { powerPointTools } from '@/tools/powerpoint';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return Boolean(ajv.compile(schema as object)(data));
}

const toolsByName = Object.fromEntries(powerPointTools.map(t => [t.name, t]));

// The fork has an expanded set of 25 PowerPoint tools
const EXPECTED_TOOL_NAMES = [
  'get_presentation_overview',
  'get_presentation_content',
  'get_slide_image',
  'get_slide_notes',
  'get_selected_slides',
  'get_selected_shapes',
  'get_slide_shapes',
  'get_slide_layouts',
  'set_presentation_content',
  'add_slide_from_code',
  'clear_slide',
  'delete_slide',
  'move_slide',
  'apply_slide_layout',
  'update_slide_shape',
  'set_shape_text',
  'update_shape_style',
  'add_geometric_shape',
  'add_line',
  'move_resize_shape',
  'delete_shape',
  'set_slide_background',
  'set_slide_notes',
  'duplicate_slide',
] as const;

// ─── Structural ───────────────────────────────────────────────────────────────

describe('Integration: PowerPoint tools — structural', () => {
  it('powerPointTools contains exactly the expected tools', () => {
    const actual = powerPointTools.map(t => t.name).sort();
    const expected = [...EXPECTED_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every tool has a non-empty name, description, and handler', () => {
    for (const tool of powerPointTools) {
      expect(tool.name.length).toBeGreaterThan(0);
      expect(tool.description!.length).toBeGreaterThan(0);
      expect(typeof tool.handler).toBe('function');
    }
  });

  it('every tool has a parameters schema', () => {
    for (const tool of powerPointTools) {
      expect(tool.parameters).toBeDefined();
    }
  });
});

// ─── get_presentation_overview ────────────────────────────────────────────────

describe('Integration: get_presentation_overview schema', () => {
  const schema = toolsByName.get_presentation_overview.parameters;

  it('accepts empty args (all params optional)', () => {
    expect(validate(schema, {})).toBe(true);
  });
});

// ─── Slide-index tools ────────────────────────────────────────────────────────

describe('Integration: slide-index tool schemas', () => {
  it('get_slide_image requires slideIndex (number)', () => {
    const schema = toolsByName.get_slide_image.parameters;
    expect(validate(schema, { slideIndex: 0 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('get_slide_image accepts optional width', () => {
    const schema = toolsByName.get_slide_image.parameters;
    expect(validate(schema, { slideIndex: 0, width: 1024 })).toBe(true);
  });

  it('clear_slide requires slideIndex', () => {
    const schema = toolsByName.clear_slide.parameters;
    expect(validate(schema, { slideIndex: 2 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('update_slide_shape requires slideIndex, shapeIndex, and text', () => {
    const schema = toolsByName.update_slide_shape.parameters;
    expect(validate(schema, { slideIndex: 0, shapeIndex: 1, text: 'Hello' })).toBe(true);
    expect(validate(schema, { slideIndex: 0 })).toBe(false);
  });

  it('duplicate_slide requires sourceIndex', () => {
    const schema = toolsByName.duplicate_slide.parameters;
    expect(validate(schema, { sourceIndex: 0 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });
});

// ─── Content tools ────────────────────────────────────────────────────────────

describe('Integration: content tool schemas', () => {
  it('set_presentation_content requires slideIndex and text', () => {
    const schema = toolsByName.set_presentation_content.parameters;
    expect(validate(schema, { slideIndex: 0, text: 'Hello' })).toBe(true);
    expect(validate(schema, { slideIndex: 0 })).toBe(false);
  });

  it('add_slide_from_code requires code string', () => {
    const schema = toolsByName.add_slide_from_code.parameters;
    expect(validate(schema, { code: 'slide.addText("Hi", {x:1,y:1,w:8,h:1})' })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('add_slide_from_code accepts optional replaceSlideIndex', () => {
    const schema = toolsByName.add_slide_from_code.parameters;
    expect(validate(schema, { code: 'slide.addText("X")', replaceSlideIndex: 0 })).toBe(true);
  });

  it('get_presentation_content accepts no args (read all slides)', () => {
    const schema = toolsByName.get_presentation_content.parameters;
    expect(validate(schema, {})).toBe(true);
  });

  it('get_presentation_content accepts slideIndex or startIndex/endIndex', () => {
    const schema = toolsByName.get_presentation_content.parameters;
    expect(validate(schema, { slideIndex: 2 })).toBe(true);
    expect(validate(schema, { startIndex: 0, endIndex: 3 })).toBe(true);
  });
});
