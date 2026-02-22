import type { Tool, ToolResultObject } from '@github/copilot-sdk';

const getDocumentOverview: Tool = {
  name: 'get_document_overview',
  description:
    'Get a structural overview of the Word document including heading hierarchy, paragraph count, tables, and content controls. Use this first to understand the document structure.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        body.load('text');

        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        for (const para of paragraphs.items) {
          para.load(['text', 'style', 'isListItem']);
        }
        await context.sync();

        const headings: string[] = [];
        let paragraphCount = 0;
        let tableCount = 0;

        for (const para of paragraphs.items) {
          paragraphCount++;
          const style = para.style ?? '';
          if (style.startsWith('Heading')) {
            const level = style.replace('Heading ', 'H');
            headings.push(`${level}: ${para.text.trim()}`);
          }
        }

        const tables = body.tables;
        tables.load('items');
        await context.sync();
        tableCount = tables.items.length;

        const summary = [
          `Document Overview`,
          `${'='.repeat(40)}`,
          `Paragraphs: ${String(paragraphCount)}`,
          `Tables: ${String(tableCount)}`,
          '',
          headings.length > 0 ? `Headings:\n${headings.join('\n')}` : '(no headings found)',
        ].join('\n');

        return summary;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getDocumentContent: Tool = {
  name: 'get_document_content',
  description:
    'Get the full HTML content of the Word document body. Returns rich formatted content that preserves structure such as headings, lists, and tables.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        const htmlResult = body.getHtml();
        await context.sync();
        return htmlResult.value;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getDocumentSection: Tool = {
  name: 'get_document_section',
  description:
    'Get the HTML content of a specific section identified by a heading. Finds the heading by partial text match and returns the content until the next heading of the same or higher level.',
  parameters: {
    type: 'object',
    properties: {
      headingText: {
        type: 'string',
        description: 'Partial or full text of the heading that starts the section.',
      },
    },
    required: ['headingText'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { headingText } = (args ?? {}) as { headingText: string };
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        for (const para of paragraphs.items) {
          para.load(['text', 'style']);
        }
        await context.sync();

        const headingPara = paragraphs.items.find(
          p =>
            p.style?.startsWith('Heading') &&
            p.text.toLowerCase().includes(headingText.toLowerCase())
        );

        if (!headingPara) {
          return `No heading found containing "${headingText}".`;
        }

        const headingLevel = parseInt(headingPara.style.replace('Heading ', ''), 10) || 1;
        const range = headingPara.getRange();

        // Expand range to include content until next heading of same or higher level
        const allParas = paragraphs.items;
        const startIdx = allParas.indexOf(headingPara);

        let endPara: Word.Paragraph | undefined;
        for (let i = startIdx + 1; i < allParas.length; i++) {
          const p = allParas[i];
          if (p.style?.startsWith('Heading')) {
            const level = parseInt(p.style.replace('Heading ', ''), 10) || 1;
            if (level <= headingLevel) {
              endPara = p;
              break;
            }
          }
        }

        let sectionRange: Word.Range;
        if (endPara) {
          const endRange = endPara.getRange(Word.RangeLocation.start);
          sectionRange = range.expandTo(endRange);
        } else {
          const bodyEnd = body.getRange(Word.RangeLocation.end);
          sectionRange = range.expandTo(bodyEnd);
        }

        const htmlResult = sectionRange.getHtml();
        await context.sync();
        return htmlResult.value;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setDocumentContent: Tool = {
  name: 'set_document_content',
  description:
    'Replace the entire document body with new HTML content. WARNING: This clears all existing content.',
  parameters: {
    type: 'object',
    properties: {
      html: {
        type: 'string',
        description: 'HTML content to set as the full document body.',
      },
    },
    required: ['html'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { html } = (args ?? {}) as { html: string };
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        body.clear();
        body.insertHtml(html, Word.InsertLocation.start);
        await context.sync();
        return 'Document content replaced successfully.';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelection: Tool = {
  name: 'get_selection',
  description:
    'Get the currently selected content in the Word document as OOXML. Useful for inspecting the structure of the selection.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const ooxmlResult = selection.getOoxml();
        await context.sync();

        const ooxml = ooxmlResult.value;
        // Extract the <w:document> element for a cleaner view
        const docMatch = /<w:document[^>]*>[\s\S]*<\/w:document>/i.exec(ooxml);
        return docMatch ? docMatch[0] : ooxml;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelectionText: Tool = {
  name: 'get_selection_text',
  description: 'Get the plain text of the currently selected content in the Word document.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        return selection.text.length > 0 ? selection.text : '(no text selected)';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertContentAtSelection: Tool = {
  name: 'insert_content_at_selection',
  description: 'Insert HTML content at or relative to the current selection in the Word document.',
  parameters: {
    type: 'object',
    properties: {
      html: { type: 'string', description: 'HTML content to insert.' },
      location: {
        type: 'string',
        enum: ['Replace', 'Before', 'After', 'Start', 'End'],
        description:
          'Where to insert relative to the selection. Replace overwrites the selection. Default: Replace.',
      },
    },
    required: ['html'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { html, location = 'Replace' } = (args ?? {}) as {
      html: string;
      location?: string;
    };

    const locationMap: Record<string, Word.InsertLocation> = {
      Replace: Word.InsertLocation.replace,
      Before: Word.InsertLocation.before,
      After: Word.InsertLocation.after,
      Start: Word.InsertLocation.start,
      End: Word.InsertLocation.end,
    };

    const insertLocation = locationMap[location] ?? Word.InsertLocation.replace;

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.insertHtml(html, insertLocation);
        await context.sync();
        return `Content inserted at selection (location: ${location}).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const findAndReplace: Tool = {
  name: 'find_and_replace',
  description: 'Find text in the Word document and replace all occurrences with new text.',
  parameters: {
    type: 'object',
    properties: {
      find: { type: 'string', description: 'Text to search for.' },
      replace: { type: 'string', description: 'Replacement text.' },
      matchCase: {
        type: 'boolean',
        description: 'Whether the search is case-sensitive. Default: false.',
      },
      matchWholeWord: {
        type: 'boolean',
        description: 'Whether to match whole words only. Default: false.',
      },
    },
    required: ['find', 'replace'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      find,
      replace,
      matchCase = false,
      matchWholeWord = false,
    } = (args ?? {}) as {
      find: string;
      replace: string;
      matchCase?: boolean;
      matchWholeWord?: boolean;
    };
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        const results = body.search(find, { matchCase, matchWholeWord });
        results.load('items');
        await context.sync();

        const count = results.items.length;
        if (count === 0) {
          return `No occurrences of "${find}" found.`;
        }

        for (const result of results.items) {
          result.insertText(replace, Word.InsertLocation.replace);
        }
        await context.sync();

        return `Replaced ${String(count)} occurrence(s) of "${find}" with "${replace}".`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertTable: Tool = {
  name: 'insert_table',
  description:
    'Insert a table at the current selection in the Word document. Supports grid, striped, and plain styles.',
  parameters: {
    type: 'object',
    properties: {
      rows: { type: 'number', description: 'Number of rows.' },
      columns: { type: 'number', description: 'Number of columns.' },
      data: {
        type: 'array',
        items: { type: 'array', items: { type: 'string' } },
        description: 'Optional 2D array of cell values (row-major). Omit for an empty table.',
      },
      style: {
        type: 'string',
        enum: ['grid', 'striped', 'plain'],
        description:
          'Visual style. grid adds borders, striped alternates row colors, plain has no formatting. Default: grid.',
      },
      hasHeaderRow: {
        type: 'boolean',
        description: 'Whether the first row is a header row with distinct styling. Default: true.',
      },
    },
    required: ['rows', 'columns'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      rows,
      columns,
      data,
      style = 'grid',
      hasHeaderRow = true,
    } = (args ?? {}) as {
      rows: number;
      columns: number;
      data?: string[][];
      style?: 'grid' | 'striped' | 'plain';
      hasHeaderRow?: boolean;
    };

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();

        // Build flat cell data array expected by insertTable
        const cellValues: string[][] = [];
        for (let r = 0; r < rows; r++) {
          const row: string[] = [];
          for (let c = 0; c < columns; c++) {
            row.push(data?.[r]?.[c] ?? '');
          }
          cellValues.push(row);
        }

        const table = selection.insertTable(rows, columns, Word.InsertLocation.after, cellValues);
        await context.sync();

        table.load('rows');
        await context.sync();

        if (style === 'grid') {
          table.style = 'Table Grid';
        }

        if (hasHeaderRow && rows > 0) {
          const headerRow = table.rows.items[0];
          headerRow.load('cells');
          await context.sync();

          for (const cell of headerRow.cells.items) {
            cell.shadingColor = '#4472C4';
            const paras = cell.body.paragraphs;
            paras.load('items');
            await context.sync();
            for (const para of paras.items) {
              para.font.bold = true;
              para.font.color = '#FFFFFF';
            }
          }
        }

        if (style === 'striped') {
          for (let r = hasHeaderRow ? 1 : 0; r < table.rows.items.length; r++) {
            if (r % 2 === 0) continue;
            const row = table.rows.items[r];
            row.load('cells');
            await context.sync();
            for (const cell of row.cells.items) {
              cell.shadingColor = '#E8E8E8';
            }
          }
        }

        await context.sync();
        return `Inserted a ${String(rows)}×${String(columns)} table with "${style}" style.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const applyStyleToSelection: Tool = {
  name: 'apply_style_to_selection',
  description: 'Apply font formatting to the currently selected text in the Word document.',
  parameters: {
    type: 'object',
    properties: {
      bold: { type: 'boolean', description: 'Make text bold.' },
      italic: { type: 'boolean', description: 'Make text italic.' },
      underline: { type: 'boolean', description: 'Underline text. true = single underline.' },
      strikeThrough: { type: 'boolean', description: 'Apply strikethrough.' },
      fontSize: { type: 'number', description: 'Font size in points.' },
      fontName: { type: 'string', description: 'Font family name (e.g. "Calibri", "Arial").' },
      fontColor: {
        type: 'string',
        description: 'Font color as hex (e.g. "#FF0000") or named color.',
      },
      highlightColor: {
        type: 'string',
        description:
          'Highlight color. Allowed values: Yellow, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, Black, White, None.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      bold,
      italic,
      underline,
      strikeThrough,
      fontSize,
      fontName,
      fontColor,
      highlightColor,
    } = (args ?? {}) as {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strikeThrough?: boolean;
      fontSize?: number;
      fontName?: string;
      fontColor?: string;
      highlightColor?: string;
    };

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const font = selection.font;

        if (bold !== undefined) font.bold = bold;
        if (italic !== undefined) font.italic = italic;
        if (underline !== undefined)
          font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        if (strikeThrough !== undefined) font.strikeThrough = strikeThrough;
        if (fontSize !== undefined) font.size = fontSize;
        if (fontName !== undefined) font.name = fontName;
        if (fontColor !== undefined) font.color = fontColor;
        if (highlightColor !== undefined) font.highlightColor = highlightColor;

        await context.sync();

        const applied: string[] = [];
        if (bold !== undefined) applied.push(`bold=${String(bold)}`);
        if (italic !== undefined) applied.push(`italic=${String(italic)}`);
        if (underline !== undefined) applied.push(`underline=${String(underline)}`);
        if (strikeThrough !== undefined) applied.push(`strikeThrough=${String(strikeThrough)}`);
        if (fontSize !== undefined) applied.push(`fontSize=${String(fontSize)}`);
        if (fontName !== undefined) applied.push(`fontName="${fontName}"`);
        if (fontColor !== undefined) applied.push(`color="${fontColor}"`);
        if (highlightColor !== undefined) applied.push(`highlight="${highlightColor}"`);

        return applied.length > 0
          ? `Applied: ${applied.join(', ')}.`
          : 'No style properties specified.';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertParagraph: Tool = {
  name: 'insert_paragraph',
  description:
    'Insert a paragraph at the end or beginning of the document body. Optionally apply a named style such as "Heading 1" or "Normal".',
  parameters: {
    type: 'object',
    properties: {
      text: { type: 'string', description: 'The paragraph text to insert.' },
      location: {
        type: 'string',
        enum: ['End', 'Start'],
        description: 'Where to insert the paragraph. Default: End.',
      },
      style: {
        type: 'string',
        description: 'Optional Word style name to apply (e.g. "Heading 1", "Normal", "Title").',
      },
    },
    required: ['text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { text, location = 'End', style } = (args ?? {}) as {
      text: string;
      location?: string;
      style?: string;
    };
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        const insertLoc =
          location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end;
        const paragraph = body.insertParagraph(text, insertLoc);
        if (style) {
          paragraph.style = style;
        }
        await context.sync();
        return `Paragraph inserted at ${location}${style ? ` with style "${style}"` : ''}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertBreak: Tool = {
  name: 'insert_break',
  description:
    'Insert a page break or section break after the current selection.',
  parameters: {
    type: 'object',
    properties: {
      breakType: {
        type: 'string',
        enum: ['page', 'sectionNext', 'sectionContinuous'],
        description: 'Type of break to insert. Default: page.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { breakType = 'page' } = (args ?? {}) as { breakType?: string };

    const breakMap: Record<string, Word.BreakType> = {
      page: Word.BreakType.page,
      sectionNext: Word.BreakType.sectionNext,
      sectionContinuous: Word.BreakType.sectionContinuous,
    };

    const wordBreakType = breakMap[breakType] ?? Word.BreakType.page;

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.insertBreak(wordBreakType, Word.InsertLocation.after);
        await context.sync();
        return `Inserted ${breakType} break after selection.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const applyParagraphStyle: Tool = {
  name: 'apply_paragraph_style',
  description:
    'Apply a named style (e.g. "Heading 1", "Title", "Normal", "Quote") to every paragraph in the current selection.',
  parameters: {
    type: 'object',
    properties: {
      styleName: {
        type: 'string',
        description: 'The Word style name to apply (e.g. "Heading 1", "Title", "Normal", "Quote").',
      },
    },
    required: ['styleName'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { styleName } = (args ?? {}) as { styleName: string };
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load('items');
        await context.sync();

        for (const para of paragraphs.items) {
          para.style = styleName;
        }
        await context.sync();

        return `Applied style "${styleName}" to ${String(paragraphs.items.length)} paragraph(s).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setParagraphFormat: Tool = {
  name: 'set_paragraph_format',
  description:
    'Set paragraph formatting on the current selection. Only provided properties are changed; others are preserved.',
  parameters: {
    type: 'object',
    properties: {
      alignment: {
        type: 'string',
        enum: ['left', 'center', 'right', 'justified'],
        description: 'Horizontal alignment.',
      },
      lineSpacing: { type: 'number', description: 'Line spacing in points.' },
      spaceBefore: { type: 'number', description: 'Space before paragraph in points.' },
      spaceAfter: { type: 'number', description: 'Space after paragraph in points.' },
      firstLineIndent: { type: 'number', description: 'First line indent in points.' },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { alignment, lineSpacing, spaceBefore, spaceAfter, firstLineIndent } = (args ?? {}) as {
      alignment?: string;
      lineSpacing?: number;
      spaceBefore?: number;
      spaceAfter?: number;
      firstLineIndent?: number;
    };

    const alignmentMap: Record<string, Word.Alignment> = {
      left: Word.Alignment.left,
      center: Word.Alignment.centered,
      right: Word.Alignment.right,
      justified: Word.Alignment.justified,
    };

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load('items');
        await context.sync();

        for (const para of paragraphs.items) {
          if (alignment !== undefined && alignmentMap[alignment] !== undefined) {
            para.alignment = alignmentMap[alignment];
          }
          if (lineSpacing !== undefined) para.lineSpacing = lineSpacing;
          if (spaceBefore !== undefined) para.spaceBefore = spaceBefore;
          if (spaceAfter !== undefined) para.spaceAfter = spaceAfter;
          if (firstLineIndent !== undefined) para.firstLineIndent = firstLineIndent;
        }
        await context.sync();

        const applied: string[] = [];
        if (alignment !== undefined) applied.push(`alignment=${alignment}`);
        if (lineSpacing !== undefined) applied.push(`lineSpacing=${String(lineSpacing)}`);
        if (spaceBefore !== undefined) applied.push(`spaceBefore=${String(spaceBefore)}`);
        if (spaceAfter !== undefined) applied.push(`spaceAfter=${String(spaceAfter)}`);
        if (firstLineIndent !== undefined)
          applied.push(`firstLineIndent=${String(firstLineIndent)}`);

        return applied.length > 0
          ? `Applied to ${String(paragraphs.items.length)} paragraph(s): ${applied.join(', ')}.`
          : 'No formatting properties specified.';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getDocumentProperties: Tool = {
  name: 'get_document_properties',
  description:
    'Get document metadata including author, title, subject, keywords, creation date, last modified time, and revision number.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const props = context.document.properties;
        props.load([
          'author',
          'title',
          'subject',
          'keywords',
          'creationDate',
          'lastSaveTime',
          'lastAuthor',
          'revisionNumber',
          'category',
          'comments',
          'company',
        ]);

        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const lines = [
          'Document Properties',
          '='.repeat(40),
          `Title: ${props.title || '(none)'}`,
          `Author: ${props.author || '(none)'}`,
          `Last Author: ${props.lastAuthor || '(none)'}`,
          `Subject: ${props.subject || '(none)'}`,
          `Keywords: ${props.keywords || '(none)'}`,
          `Category: ${props.category || '(none)'}`,
          `Company: ${props.company || '(none)'}`,
          `Comments: ${props.comments || '(none)'}`,
          `Creation Date: ${String(props.creationDate)}`,
          `Last Save Time: ${String(props.lastSaveTime)}`,
          `Revision Number: ${props.revisionNumber || '(none)'}`,
          `Paragraph Count: ${String(paragraphs.items.length)}`,
        ];

        return lines.join('\n');
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertImage: Tool = {
  name: 'insert_image',
  description:
    'Insert a base64-encoded image at the current selection. Optionally set width and height in points.',
  parameters: {
    type: 'object',
    properties: {
      base64Image: {
        type: 'string',
        description: 'Base64-encoded image data (without data URI prefix).',
      },
      width: { type: 'number', description: 'Image width in points.' },
      height: { type: 'number', description: 'Image height in points.' },
    },
    required: ['base64Image'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { base64Image, width, height } = (args ?? {}) as {
      base64Image: string;
      width?: number;
      height?: number;
    };
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const picture = selection.insertInlinePictureFromBase64(
          base64Image,
          Word.InsertLocation.replace
        );
        if (width !== undefined) picture.width = width;
        if (height !== undefined) picture.height = height;
        await context.sync();
        return `Image inserted${width ? ` (width=${String(width)})` : ''}${height ? ` (height=${String(height)})` : ''}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getComments: Tool = {
  name: 'get_comments',
  description:
    'Get all comments from the document body, including author, content, creation date, and resolved status.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const comments = context.document.body.getComments();
        comments.load('items');
        await context.sync();

        if (comments.items.length === 0) {
          return '(no comments found)';
        }

        for (const comment of comments.items) {
          comment.load(['authorName', 'authorEmail', 'content', 'creationDate', 'resolved', 'id']);
        }
        await context.sync();

        const lines = comments.items.map(
          (c, i) =>
            `${String(i + 1)}. [${c.resolved ? 'Resolved' : 'Open'}] ${c.authorName} (${c.authorEmail}): "${c.content}" — ${String(c.creationDate)}`
        );

        return `Comments (${String(comments.items.length)}):\n${lines.join('\n')}`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertList: Tool = {
  name: 'insert_list',
  description:
    'Insert a bulleted or numbered list at the current selection using HTML. Provide the list items as a newline-separated string.',
  parameters: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: 'List items separated by newlines. Each line becomes a list item.',
      },
      listType: {
        type: 'string',
        enum: ['bullet', 'number'],
        description: 'Type of list to insert. Default: bullet.',
      },
    },
    required: ['text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { text, listType = 'bullet' } = (args ?? {}) as {
      text: string;
      listType?: string;
    };
    try {
      return await Word.run(async context => {
        const items = text
          .split('\n')
          .map(line => line.trim())
          .filter(Boolean);
        const tag = listType === 'number' ? 'ol' : 'ul';
        const html = `<${tag}>${items.map(item => `<li>${item}</li>`).join('')}</${tag}>`;

        const selection = context.document.getSelection();
        selection.insertHtml(html, Word.InsertLocation.replace);
        await context.sync();

        return `Inserted ${listType === 'number' ? 'numbered' : 'bulleted'} list with ${String(items.length)} item(s).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getContentControls: Tool = {
  name: 'get_content_controls',
  description:
    'List all content controls in the document, including their tag, title, text, and type.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const controls = context.document.contentControls;
        controls.load('items');
        await context.sync();

        if (controls.items.length === 0) {
          return '(no content controls found)';
        }

        for (const cc of controls.items) {
          cc.load(['tag', 'title', 'text', 'type']);
        }
        await context.sync();

        const lines = controls.items.map(
          (cc, i) =>
            `${String(i + 1)}. [${String(cc.type)}] tag="${cc.tag}" title="${cc.title}" text="${cc.text.length > 100 ? cc.text.slice(0, 100) + '…' : cc.text}"`
        );

        return `Content Controls (${String(controls.items.length)}):\n${lines.join('\n')}`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertTextAtBookmark: Tool = {
  name: 'insert_text_at_bookmark',
  description:
    'Insert text at a named bookmark location in the document. Can replace the bookmark content, or insert before or after it.',
  parameters: {
    type: 'object',
    properties: {
      bookmarkName: {
        type: 'string',
        description: 'The name of the bookmark (case-insensitive).',
      },
      text: { type: 'string', description: 'Text to insert.' },
      insertLocation: {
        type: 'string',
        enum: ['Before', 'After', 'Replace'],
        description: 'Where to insert relative to the bookmark. Default: Replace.',
      },
    },
    required: ['bookmarkName', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { bookmarkName, text, insertLocation = 'Replace' } = (args ?? {}) as {
      bookmarkName: string;
      text: string;
      insertLocation?: string;
    };

    const locationMap: Record<string, Word.InsertLocation> = {
      Before: Word.InsertLocation.before,
      After: Word.InsertLocation.after,
      Replace: Word.InsertLocation.replace,
    };

    const loc = locationMap[insertLocation] ?? Word.InsertLocation.replace;

    try {
      return await Word.run(async context => {
        const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
        range.load('isNullObject');
        await context.sync();

        if (range.isNullObject) {
          return `Bookmark "${bookmarkName}" not found.`;
        }

        range.insertText(text, loc);
        await context.sync();

        return `Text inserted at bookmark "${bookmarkName}" (location: ${insertLocation}).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

export const wordTools: Tool[] = [
  getDocumentOverview,
  getDocumentContent,
  getDocumentSection,
  setDocumentContent,
  getSelection,
  getSelectionText,
  insertContentAtSelection,
  findAndReplace,
  insertTable,
  applyStyleToSelection,
  insertParagraph,
  insertBreak,
  applyParagraphStyle,
  setParagraphFormat,
  getDocumentProperties,
  insertImage,
  getComments,
  insertList,
  getContentControls,
  insertTextAtBookmark,
];
