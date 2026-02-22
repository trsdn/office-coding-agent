import { describe, it, expect } from 'vitest';

/**
 * Tests for the Word document orchestrator trigger detection.
 * Three regex patterns determine when to activate the orchestrator:
 * 1. Deep keywords (gründlich, deep, think, etc.)
 * 2. Multi-section requests ("5 sections", "3 Kapitel")
 * 3. Document creation requests ("write a report", "erstelle ein Dokument")
 */

const DEEP_KEYWORD = /\b(deep|gründlich|ausführlich|thoroughly|think|go\s*deep|detail(liert)?|qualit)/i;
const MULTI_SECTION = /\b(\d+)\s*(sections?|abschnitt(e|en)?|kapitel|teil(e|en)?|chapters?)\b/i;
const DOC_CREATION = /\b(erstell|schreib|create|write|build|generate|verfass)\w*\b.{0,30}\b(report|bericht|dokument|document|paper|aufsatz|memo|proposal|angebot|zusammenfassung)\b/i;

function shouldTriggerOrchestrator(input: string): boolean {
  return DEEP_KEYWORD.test(input) || MULTI_SECTION.test(input) || DOC_CREATION.test(input);
}

describe('Word orchestrator trigger — deep keywords', () => {
  const shouldTrigger = [
    'Schreibe einen Report gründlich',
    'Create a document deep',
    'Go deep on this topic',
    'Think about this and write it',
    'Write a thoroughly researched paper',
    'Erstelle ein ausführliches Dokument',
    'Write a detailed report',
    'Erstelle einen detaillierten Bericht',
    'High quality document please',
  ];

  it.each(shouldTrigger)('triggers for: "%s"', (input) => {
    expect(DEEP_KEYWORD.test(input)).toBe(true);
  });
});

describe('Word orchestrator trigger — multi-section requests', () => {
  const shouldTrigger = [
    'Create a document with 5 sections',
    'Schreibe 3 Abschnitte',
    'Write 4 chapters about AI',
    'Erstelle 6 Kapitel',
    'Build a report with 2 sections',
    'Dokument mit 3 Teilen',
  ];

  const shouldNotTrigger = [
    'Summarize this document',
    'Insert a table',
    'Format the heading',
  ];

  it.each(shouldTrigger)('triggers for: "%s"', (input) => {
    expect(MULTI_SECTION.test(input)).toBe(true);
  });

  it.each(shouldNotTrigger)('does NOT trigger for: "%s"', (input) => {
    expect(MULTI_SECTION.test(input)).toBe(false);
  });
});

describe('Word orchestrator trigger — document creation', () => {
  const shouldTrigger = [
    'Erstelle einen Bericht über KI',
    'Write a report about sales',
    'Schreibe ein Dokument über das Projekt',
    'Create a document about our strategy',
    'Generate a memo for the team',
    'Verfasse eine Zusammenfassung',
    'Build a proposal for the client',
  ];

  const shouldNotTrigger = [
    'Summarize this document',
    'Add a paragraph',
    'Find and replace text',
    'Get the document structure',
  ];

  it.each(shouldTrigger)('triggers for: "%s"', (input) => {
    expect(DOC_CREATION.test(input)).toBe(true);
  });

  it.each(shouldNotTrigger)('does NOT trigger for: "%s"', (input) => {
    expect(DOC_CREATION.test(input)).toBe(false);
  });
});

describe('Word orchestrator — combined trigger', () => {
  it.each([
    ['Write a report about AI', true],
    ['Erstelle 5 Kapitel', true],
    ['Go deep on this', true],
    ['Insert a table', false],
    ['Format the heading', false],
  ])('"%s" → orchestrator: %s', (input, expected) => {
    expect(shouldTriggerOrchestrator(input)).toBe(expected);
  });
});

describe('Word deep-mode selection', () => {
  const DEEP_MODE_REGEX = /\b(deep|gründlich|ausführlich|thoroughly|think|go\s*deep|detail(liert)?|qualit)/i;

  it.each([
    ['deep', 'deep'],
    ['gründlich', 'deep'],
    ['detailliert', 'deep'],
    ['quality report', 'deep'],
    ['Write a report about AI', 'fast'],
    ['Create 5 sections', 'fast'],
  ])('"%s" → %s mode', (input, expected) => {
    const mode = DEEP_MODE_REGEX.test(input) ? 'deep' : 'fast';
    expect(mode).toBe(expected);
  });
});
