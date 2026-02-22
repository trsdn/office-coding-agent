import type { Tool, ToolResultObject } from '@github/copilot-sdk';
import pptxgen from 'pptxgenjs';

const CHUNK_SIZE = 10;

async function getSlideCount(): Promise<number> {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    return slides.items.length;
  });
}

async function getSlideTextContent(startIdx: number, endIdx: number): Promise<string[]> {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    const slideRefs = slides.items.slice(startIdx, endIdx + 1);
    for (const slide of slideRefs) {
      slide.shapes.load('items');
    }
    await context.sync();

    const results: string[] = [];
    // Process each slide individually so one bad shape doesn't break everything
    for (let i = 0; i < slideRefs.length; i++) {
      const slide = slideRefs[i];
      const texts: string[] = [];
      let shapeCount = 0;

      // Process each shape individually — batch sync fails when ANY shape lacks a textFrame
      for (const shape of slide.shapes.items) {
        shapeCount++;
        try {
          shape.textFrame.textRange.load('text');
          await context.sync();
          const text = shape.textFrame.textRange.text?.trim() ?? '';
          if (text.length > 0) texts.push(text);
        } catch {
          // shape doesn't support textFrame (SmartArt, chart, image, etc.)
        }
      }

      if (texts.length > 0) {
        results.push(`Slide ${startIdx + i + 1}: ${texts.join(' | ')}`);
      } else if (shapeCount > 0) {
        results.push(
          `Slide ${startIdx + i + 1}: (contains graphics/SmartArt — use get_slide_image to see visual content)`
        );
      } else {
        results.push(`Slide ${startIdx + i + 1}: (empty slide)`);
      }
    }
    return results;
  });
}

const getPresentationOverview: Tool = {
  name: 'get_presentation_overview',
  description:
    "Get an overview of the entire PowerPoint presentation including total slide count and a text preview of each slide's shapes. Use this first to understand the presentation structure.",
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const slideCount = await getSlideCount();
      if (slideCount === 0) return 'Presentation has no slides.';

      const slideOverviews: string[] = [];
      for (let i = 0; i < slideCount; i += CHUNK_SIZE) {
        const chunkEnd = Math.min(i + CHUNK_SIZE - 1, slideCount - 1);
        const chunk = await getSlideTextContent(i, chunkEnd);
        slideOverviews.push(...chunk);
      }

      return `Presentation Overview\n${'='.repeat(40)}\nTotal slides: ${String(slideCount)}\n\n${slideOverviews.join('\n')}`;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getPresentationContent: Tool = {
  name: 'get_presentation_content',
  description:
    'Get the text content of one or more slides. Specify slideIndex for a single slide, startIndex/endIndex for a range, or omit all to read every slide.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '0-based index for reading a single slide. Omit to use range or read all.',
      },
      startIndex: { type: 'number', description: '0-based start index for a range of slides.' },
      endIndex: {
        type: 'number',
        description: '0-based end index (inclusive) for a range of slides.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, startIndex, endIndex } = (args ?? {}) as {
      slideIndex?: number;
      startIndex?: number;
      endIndex?: number;
    };
    try {
      const slideCount = await getSlideCount();
      if (slideCount === 0) return 'Presentation has no slides.';

      let start: number;
      let end: number;

      if (slideIndex !== undefined) {
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }
        start = slideIndex;
        end = slideIndex;
      } else if (startIndex !== undefined && endIndex !== undefined) {
        start = Math.max(0, startIndex);
        end = Math.min(slideCount - 1, endIndex);
        if (start > end) {
          return {
            textResultForLlm: `Invalid range: startIndex (${String(startIndex)}) must be <= endIndex (${String(endIndex)}).`,
            resultType: 'failure',
            error: 'Invalid range',
            toolTelemetry: {},
          };
        }
      } else {
        start = 0;
        end = slideCount - 1;
      }

      const results: string[] = [];
      for (let i = start; i <= end; i += CHUNK_SIZE) {
        const chunkEnd = Math.min(i + CHUNK_SIZE - 1, end);
        const chunk = await getSlideTextContent(i, chunkEnd);
        results.push(...chunk);
      }
      return results.join('\n');
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

// Crop a base64 PNG to a region using OffscreenCanvas (available in modern browsers)
function cropImage(
  base64: string,
  region: 'full' | 'top' | 'bottom' | 'left' | 'right',
): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const w = img.width;
      const h = img.height;
      let sx = 0,
        sy = 0,
        sw = w,
        sh = h;
      switch (region) {
        case 'top':
          sh = Math.round(h * 0.5);
          break;
        case 'bottom':
          sy = Math.round(h * 0.5);
          sh = h - sy;
          break;
        case 'left':
          sw = Math.round(w * 0.5);
          break;
        case 'right':
          sx = Math.round(w * 0.5);
          sw = w - sx;
          break;
      }
      const canvas = document.createElement('canvas');
      canvas.width = sw;
      canvas.height = sh;
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        reject(new Error('Canvas 2D context unavailable'));
        return;
      }
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => reject(new Error('Failed to load image for cropping'));
    img.src = `data:image/png;base64,${base64}`;
  });
}

const getSlideImage: Tool = {
  name: 'get_slide_image',
  description:
    'Capture a slide (or region) as a PNG image to verify visual quality. Use region="bottom" to zoom into the bottom half where text overflow usually occurs. Returns a base64 data URI.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      region: {
        type: 'string',
        enum: ['full', 'top', 'bottom', 'left', 'right'],
        description:
          'Which part of the slide to capture. "full" = entire slide (default). "bottom" = bottom half (best for checking text overflow). "top"/"left"/"right" for other regions.',
      },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, region = 'full' } = (args ?? {}) as {
      slideIndex: number;
      region?: 'full' | 'top' | 'bottom' | 'left' | 'right';
    };
    // Use higher resolution when cropping a region (half the data since we crop 50%)
    const captureWidth = region === 'full' ? 600 : 960;
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        const imageResult = slide.getImageAsBase64({ width: captureWidth });
        await context.sync();

        const base64 = imageResult.value;

        // Crop to requested region
        let dataUri: string;
        if (region !== 'full') {
          dataUri = await cropImage(base64, region);
        } else {
          dataUri = `data:image/png;base64,${base64}`;
        }

        // If still too large after cropping, re-capture at smaller width
        if (dataUri.length > 150000) {
          const smallImage = slide.getImageAsBase64({ width: 400 });
          await context.sync();
          let smallUri: string;
          if (region !== 'full') {
            smallUri = await cropImage(smallImage.value, region);
          } else {
            smallUri = `data:image/png;base64,${smallImage.value}`;
          }
          return {
            textResultForLlm: `[Auto-reduced to 400px because image was ${String(Math.round(dataUri.length / 1024))}KB — region: ${region}]\n${smallUri}`,
            resultType: 'success',
            toolTelemetry: {},
          };
        }

        const label = region === 'full' ? '' : ` [region: ${region}]`;
        return {
          textResultForLlm: `${dataUri}${label}`,
          resultType: 'success',
          toolTelemetry: {},
        };
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return {
        textResultForLlm: `Slide image capture failed: ${msg}. Ensure you are using PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web.`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }
  },
};

const getSlideNotes: Tool = {
  name: 'get_slide_notes',
  description:
    'Get speaker notes from a PowerPoint slide. Note: The notes API has limited support in web add-ins; notes may not be available in all environments.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '0-based slide index. Omit to get notes from all slides.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex?: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) return 'Presentation has no slides.';

        if (slideIndex !== undefined && (slideIndex < 0 || slideIndex >= slideCount)) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const startIdx = slideIndex ?? 0;
        const endIdx = slideIndex !== undefined ? slideIndex + 1 : slideCount;

        const results: string[] = [];
        for (let i = startIdx; i < endIdx; i++) {
          results.push(
            `Slide ${i + 1}: (Speaker notes require PowerPoint desktop — API limitation in web add-ins)`
          );
        }

        return slideIndex !== undefined
          ? results[0]
          : `Speaker Notes\n${'='.repeat(40)}\n${results.join('\n')}`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setPresentationContent: Tool = {
  name: 'set_presentation_content',
  description:
    'Add a text box to a slide. Pass slideIndex equal to the current total slide count to add a new slide first.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description:
          '0-based slide index. Pass the total slide count to append a new slide before adding.',
      },
      text: { type: 'string', description: 'The text content to add.' },
    },
    required: ['slideIndex', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, text } = (args ?? {}) as { slideIndex: number; text: string };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        let slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex > slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        if (slideIndex === slideCount) {
          context.presentation.slides.add();
          await context.sync();
          slides.load('items');
          await context.sync();
          slideCount = slides.items.length;
        }

        const slide = slides.items[slideIndex];
        slide.shapes.addTextBox(text, { left: 50, top: 100, width: 600, height: 400 });
        await context.sync();

        return `Added text box to slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const addSlideFromCode: Tool = {
  name: 'add_slide_from_code',
  description: `Add a richly formatted slide to the presentation using PptxGenJS code.
Provide a JavaScript function body that receives a 'slide' parameter (PptxGenJS Slide object).

PptxGenJS API reference:
- Text:   slide.addText("Hello", { x:1, y:1, w:8, h:1, fontSize:24, bold:true, color:"363636" })
- Bullets: slide.addText([{text:"Point 1",options:{bullet:true}},{text:"Point 2",options:{bullet:true}}], { x:0.5, y:1.5, w:9, h:3, fontSize:18 })
- Image (base64): slide.addImage({ data:"data:image/png;base64,...", x:1, y:1, w:4, h:3 })
- Table:  slide.addTable([["H1","H2"],["R1","R2"]], { x:0.5, y:2, w:9, fontSize:14 })
- Shape:  slide.addShape("rect", { x:1, y:1, w:3, h:1, fill:{ color:"FF0000" } })
- All positions (x, y, w, h) are in inches.`,
  parameters: {
    type: 'object',
    properties: {
      code: {
        type: 'string',
        description:
          "JavaScript code (function body) receiving a 'slide' parameter. Call PptxGenJS methods on it to build slide content.",
      },
      replaceSlideIndex: {
        type: 'number',
        description:
          'Optional 0-based index of an existing slide to replace. If omitted, the new slide is appended.',
      },
    },
    required: ['code'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { code, replaceSlideIndex } = (args ?? {}) as {
      code: string;
      replaceSlideIndex?: number;
    };

    const pptx = new pptxgen();
    const slide = pptx.addSlide();

    try {
      /* eslint-disable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */
      const buildSlide = new Function('slide', code);
      buildSlide(slide);
      /* eslint-enable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */
    } catch (codeError) {
      const msg = codeError instanceof Error ? codeError.message : String(codeError);
      return {
        textResultForLlm: `Code execution error: ${msg}`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }

    let base64: string;
    try {
      base64 = (await pptx.write({ outputType: 'base64' })) as string;
    } catch (writeError) {
      const msg = writeError instanceof Error ? writeError.message : String(writeError);
      return {
        textResultForLlm: `Failed to generate slide: ${msg}`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }

    try {
      await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;

        const insertOptions: PowerPoint.InsertSlideOptions = {
          formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
        };

        if (replaceSlideIndex !== undefined) {
          if (replaceSlideIndex < 0 || replaceSlideIndex >= slideCount) {
            throw new Error(
              `Invalid replaceSlideIndex ${String(replaceSlideIndex)}. Must be 0-${String(slideCount - 1)}.`
            );
          }
          if (replaceSlideIndex > 0) {
            const prevSlide = slides.items[replaceSlideIndex - 1];
            prevSlide.load('id');
            await context.sync();
            insertOptions.targetSlideId = prevSlide.id;
          }
        } else if (slideCount > 0) {
          const lastSlide = slides.items[slideCount - 1];
          lastSlide.load('id');
          await context.sync();
          insertOptions.targetSlideId = lastSlide.id;
        }

        context.presentation.insertSlidesFromBase64(base64, insertOptions);
        await context.sync();

        if (replaceSlideIndex !== undefined) {
          // After insert the old slide has shifted by 1
          slides.load('items');
          await context.sync();
          const oldSlide = slides.items[replaceSlideIndex + 1];
          oldSlide.delete();
          await context.sync();
        }
      });

      return replaceSlideIndex !== undefined
        ? `Successfully replaced slide ${String(replaceSlideIndex + 1)}.`
        : 'Successfully added new slide to the presentation.';
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const clearSlide: Tool = {
  name: 'clear_slide',
  description: 'Remove all shapes from a specific slide, leaving it blank.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();

        for (const shape of shapes.items) {
          shape.delete();
        }
        await context.sync();

        return `Cleared all shapes from slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const updateSlideShape: Tool = {
  name: 'update_slide_shape',
  description: 'Update the text content of an existing shape on a PowerPoint slide.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index within the slide.' },
      text: { type: 'string', description: 'The new text content for the shape.' },
    },
    required: ['slideIndex', 'shapeIndex', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, shapeIndex, text } = (args ?? {}) as {
      slideIndex: number;
      shapeIndex: number;
      text: string;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();

        const shapeCount = shapes.items.length;
        if (shapeIndex < 0 || shapeIndex >= shapeCount) {
          return {
            textResultForLlm: `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`,
            resultType: 'failure',
            error: 'Invalid shapeIndex',
            toolTelemetry: {},
          };
        }

        const shape = shapes.items[shapeIndex];
        shape.textFrame.textRange.load('text');
        await context.sync();

        shape.textFrame.textRange.text = text;
        await context.sync();

        return `Updated shape ${String(shapeIndex + 1)} on slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setSlideNotes: Tool = {
  name: 'set_slide_notes',
  description:
    'Add or update speaker notes for a slide. Due to API limitations, this provides guidance rather than directly modifying notes in all environments.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      notes: { type: 'string', description: 'The speaker notes text.' },
    },
    required: ['slideIndex', 'notes'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, notes } = (args ?? {}) as { slideIndex: number; notes: string };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const preview = notes.length > 100 ? `${notes.substring(0, 100)}...` : notes;
        return `Note: Direct speaker notes editing has limited API support in web add-ins. For slide ${String(slideIndex + 1)}, please use the Notes pane in PowerPoint to add: "${preview}"`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const duplicateSlide: Tool = {
  name: 'duplicate_slide',
  description:
    'Duplicate an existing slide by copying its text content into a new slide. Note: Only text shapes are copied; complex graphics may not be preserved.',
  parameters: {
    type: 'object',
    properties: {
      sourceIndex: { type: 'number', description: '0-based index of the slide to duplicate.' },
    },
    required: ['sourceIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { sourceIndex } = (args ?? {}) as { sourceIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) return 'Presentation has no slides.';
        if (sourceIndex < 0 || sourceIndex >= slideCount) {
          return `Invalid sourceIndex ${String(sourceIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        // Collect shapes from source slide
        const sourceSlide = slides.items[sourceIndex];
        sourceSlide.shapes.load('items');
        await context.sync();

        for (const shape of sourceSlide.shapes.items) {
          try {
            shape.textFrame.textRange.load('text');
          } catch {
            // not all shapes have a textFrame
          }
        }
        await context.sync();

        // Add a new blank slide and copy text shapes
        slides.add();
        await context.sync();
        slides.load('items');
        await context.sync();
        const newSlide = slides.items[slides.items.length - 1];

        for (const shape of sourceSlide.shapes.items) {
          try {
            const text = shape.textFrame.textRange.text ?? '';
            if (text) {
              newSlide.shapes.addTextBox(text, { left: 50, top: 100, width: 600, height: 100 });
            }
          } catch {
            // skip non-text shapes
          }
        }
        await context.sync();

        return `Duplicated slide ${String(sourceIndex + 1)} (text content only).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSlideShapes: Tool = {
  name: 'get_slide_shapes',
  description:
    'List all shapes on a slide with their index, type, position, size, text, and fill color. Use this to understand the layout before modifying individual shapes.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();

        if (shapes.items.length === 0) return `Slide ${String(slideIndex + 1)} has no shapes.`;

        const lines = [
          `Shapes on Slide ${String(slideIndex + 1)} (${String(shapes.items.length)} total)`,
          '='.repeat(50),
        ];

        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          // Load properties individually to avoid batch failures
          try {
            shape.load('name,left,top,width,height');
            await context.sync();
          } catch {
            lines.push(`\n${String(i)}. (unable to read shape properties)`);
            continue;
          }

          const info: string[] = [
            `\n${String(i)}. "${shape.name}"`,
            `   Position: left=${String(Math.round(shape.left))}, top=${String(Math.round(shape.top))}`,
            `   Size: width=${String(Math.round(shape.width))}, height=${String(Math.round(shape.height))}`,
          ];

          // Try to read text
          try {
            shape.textFrame.textRange.load('text');
            await context.sync();
            const text = shape.textFrame.textRange.text?.trim() ?? '';
            if (text.length > 0) {
              const preview = text.length > 80 ? `${text.substring(0, 80)}...` : text;
              info.push(`   Text: "${preview}"`);
            }
          } catch {
            // no textFrame
          }

          // Try to read fill color
          try {
            shape.fill.load('foregroundColor,type');
            await context.sync();
            if (shape.fill.foregroundColor) {
              info.push(`   Fill: ${shape.fill.foregroundColor}`);
            }
          } catch {
            // no fill info
          }

          lines.push(...info);
        }

        return lines.join('\n');
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const updateShapeStyle: Tool = {
  name: 'update_shape_style',
  description:
    'Change the visual style of an existing shape: fill color, outline color, outline width, font color, font size, bold, italic. Only specified properties are changed — others are preserved.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (use get_slide_shapes to find it).' },
      fillColor: {
        type: 'string',
        description: 'Fill color as hex (e.g. "FF0000" for red, "4472C4" for blue). Use "transparent" to remove fill.',
      },
      outlineColor: { type: 'string', description: 'Outline color as hex.' },
      outlineWidth: { type: 'number', description: 'Outline width in points.' },
      fontColor: { type: 'string', description: 'Font color as hex for all text in the shape.' },
      fontSize: { type: 'number', description: 'Font size in points for all text in the shape.' },
      bold: { type: 'boolean', description: 'Set text bold.' },
      italic: { type: 'boolean', description: 'Set text italic.' },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const a = (args ?? {}) as {
      slideIndex: number;
      shapeIndex: number;
      fillColor?: string;
      outlineColor?: string;
      outlineWidth?: number;
      fontColor?: string;
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (a.slideIndex < 0 || a.slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(a.slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const slide = slides.items[a.slideIndex];
        slide.shapes.load('items');
        await context.sync();

        const shapeCount = slide.shapes.items.length;
        if (a.shapeIndex < 0 || a.shapeIndex >= shapeCount) {
          return `Invalid shapeIndex ${String(a.shapeIndex)}. Slide has ${String(shapeCount)} shapes.`;
        }

        const shape = slide.shapes.items[a.shapeIndex];
        const changes: string[] = [];

        // Fill color
        if (a.fillColor !== undefined) {
          if (a.fillColor === 'transparent') {
            shape.fill.clear();
            changes.push('fill cleared');
          } else {
            shape.fill.setSolidColor(a.fillColor);
            changes.push(`fill=${a.fillColor}`);
          }
        }

        // Outline
        if (a.outlineColor !== undefined) {
          shape.lineFormat.color = a.outlineColor;
          changes.push(`outline color=${a.outlineColor}`);
        }
        if (a.outlineWidth !== undefined) {
          shape.lineFormat.weight = a.outlineWidth;
          changes.push(`outline width=${String(a.outlineWidth)}pt`);
        }

        // Font properties
        if (a.fontColor !== undefined || a.fontSize !== undefined || a.bold !== undefined || a.italic !== undefined) {
          try {
            const font = shape.textFrame.textRange.font;
            if (a.fontColor !== undefined) {
              font.color = a.fontColor;
              changes.push(`font color=${a.fontColor}`);
            }
            if (a.fontSize !== undefined) {
              font.size = a.fontSize;
              changes.push(`font size=${String(a.fontSize)}pt`);
            }
            if (a.bold !== undefined) {
              font.bold = a.bold;
              changes.push(`bold=${String(a.bold)}`);
            }
            if (a.italic !== undefined) {
              font.italic = a.italic;
              changes.push(`italic=${String(a.italic)}`);
            }
          } catch {
            changes.push('(font changes skipped — shape has no text)');
          }
        }

        await context.sync();

        return `Updated shape ${String(a.shapeIndex)} on slide ${String(a.slideIndex + 1)}: ${changes.join(', ')}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const moveResizeShape: Tool = {
  name: 'move_resize_shape',
  description:
    'Move and/or resize an existing shape on a slide. Only specified properties are changed — others are preserved. Coordinates are in points (1 inch = 72 points).',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (use get_slide_shapes to find it).' },
      left: { type: 'number', description: 'New left position in points.' },
      top: { type: 'number', description: 'New top position in points.' },
      width: { type: 'number', description: 'New width in points.' },
      height: { type: 'number', description: 'New height in points.' },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const a = (args ?? {}) as {
      slideIndex: number;
      shapeIndex: number;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (a.slideIndex < 0 || a.slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(a.slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const slide = slides.items[a.slideIndex];
        slide.shapes.load('items');
        await context.sync();

        const shapeCount = slide.shapes.items.length;
        if (a.shapeIndex < 0 || a.shapeIndex >= shapeCount) {
          return `Invalid shapeIndex ${String(a.shapeIndex)}. Slide has ${String(shapeCount)} shapes.`;
        }

        const shape = slide.shapes.items[a.shapeIndex];
        const changes: string[] = [];

        if (a.left !== undefined) {
          shape.left = a.left;
          changes.push(`left=${String(a.left)}`);
        }
        if (a.top !== undefined) {
          shape.top = a.top;
          changes.push(`top=${String(a.top)}`);
        }
        if (a.width !== undefined) {
          shape.width = a.width;
          changes.push(`width=${String(a.width)}`);
        }
        if (a.height !== undefined) {
          shape.height = a.height;
          changes.push(`height=${String(a.height)}`);
        }

        await context.sync();

        return changes.length > 0
          ? `Moved/resized shape ${String(a.shapeIndex)} on slide ${String(a.slideIndex + 1)}: ${changes.join(', ')}.`
          : 'No changes specified.';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const deleteSlide: Tool = {
  name: 'delete_slide',
  description: 'Delete a slide from the presentation by its 0-based index.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index to delete.' },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        slides.items[slideIndex].delete();
        await context.sync();
        return `Deleted slide ${String(slideIndex + 1)}. Presentation now has ${String(slideCount - 1)} slides.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const moveSlide: Tool = {
  name: 'move_slide',
  description: 'Move a slide to a new position in the presentation. Both indices are 0-based.',
  parameters: {
    type: 'object',
    properties: {
      fromIndex: { type: 'number', description: '0-based index of the slide to move.' },
      toIndex: { type: 'number', description: '0-based destination index.' },
    },
    required: ['fromIndex', 'toIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { fromIndex, toIndex } = (args ?? {}) as { fromIndex: number; toIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (fromIndex < 0 || fromIndex >= slideCount) {
          return `Invalid fromIndex ${String(fromIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        if (toIndex < 0 || toIndex >= slideCount) {
          return `Invalid toIndex ${String(toIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        slides.items[fromIndex].moveTo(toIndex);
        await context.sync();
        return `Moved slide from position ${String(fromIndex + 1)} to position ${String(toIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setSlideBackground: Tool = {
  name: 'set_slide_background',
  description: 'Set the background color of a slide. Use hex color values (e.g. "FFFFFF" for white, "1F4E79" for dark blue).',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      color: { type: 'string', description: 'Background color as hex (e.g. "FFFFFF", "1F4E79").' },
    },
    required: ['slideIndex', 'color'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, color } = (args ?? {}) as { slideIndex: number; color: string };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        const slide = slides.items[slideIndex];
        slide.background.fill.setSolidFill({ color });
        await context.sync();
        return `Set background of slide ${String(slideIndex + 1)} to #${color}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const addGeometricShape: Tool = {
  name: 'add_geometric_shape',
  description:
    'Add a geometric shape (rectangle, ellipse, triangle, arrow, etc.) to a slide. Position and size are in points (1 inch = 72 points). Standard slide is 720×540 points.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      geometryType: {
        type: 'string',
        description:
          'Shape type: rectangle, roundedRectangle, ellipse, triangle, rightTriangle, diamond, pentagon, hexagon, star4, star5, arrow, leftArrow, upArrow, downArrow, parallelogram, trapezoid, cloud, heart, plus, donut, noSmoking, blockArc, foldedCorner, smileyFace, and more.',
      },
      left: { type: 'number', description: 'Left position in points. Default: 100.' },
      top: { type: 'number', description: 'Top position in points. Default: 100.' },
      width: { type: 'number', description: 'Width in points. Default: 200.' },
      height: { type: 'number', description: 'Height in points. Default: 150.' },
      fillColor: { type: 'string', description: 'Fill color as hex (e.g. "4472C4").' },
    },
    required: ['slideIndex', 'geometryType'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const a = (args ?? {}) as {
      slideIndex: number;
      geometryType: string;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
      fillColor?: string;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (a.slideIndex < 0 || a.slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(a.slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        const slide = slides.items[a.slideIndex];
        const shape = slide.shapes.addGeometricShape(
          a.geometryType as PowerPoint.GeometricShapeType,
          {
            left: a.left ?? 100,
            top: a.top ?? 100,
            width: a.width ?? 200,
            height: a.height ?? 150,
          }
        );
        if (a.fillColor) {
          shape.fill.setSolidColor(a.fillColor);
        }
        await context.sync();
        return `Added ${a.geometryType} shape to slide ${String(a.slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const addLine: Tool = {
  name: 'add_line',
  description: 'Add a line or connector to a slide. Position and size are in points.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      connectorType: {
        type: 'string',
        description: 'Line type: straight, elbow, or curve. Default: straight.',
      },
      left: { type: 'number', description: 'Left (start X) in points. Default: 100.' },
      top: { type: 'number', description: 'Top (start Y) in points. Default: 100.' },
      width: { type: 'number', description: 'Horizontal length in points. Default: 200.' },
      height: { type: 'number', description: 'Vertical length in points. Default: 0 (horizontal line).' },
      lineColor: { type: 'string', description: 'Line color as hex.' },
      lineWeight: { type: 'number', description: 'Line weight in points.' },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const a = (args ?? {}) as {
      slideIndex: number;
      connectorType?: string;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
      lineColor?: string;
      lineWeight?: number;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (a.slideIndex < 0 || a.slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(a.slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        const slide = slides.items[a.slideIndex];
        const connType = (a.connectorType ?? 'straight') as PowerPoint.ConnectorType;
        const line = slide.shapes.addLine(connType, {
          left: a.left ?? 100,
          top: a.top ?? 100,
          width: a.width ?? 200,
          height: a.height ?? 0,
        });
        if (a.lineColor) {
          line.lineFormat.color = a.lineColor;
        }
        if (a.lineWeight !== undefined) {
          line.lineFormat.weight = a.lineWeight;
        }
        await context.sync();
        return `Added ${connType} line to slide ${String(a.slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const deleteShape: Tool = {
  name: 'delete_shape',
  description: 'Delete a specific shape from a slide by its 0-based index.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (use get_slide_shapes to find it).' },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, shapeIndex } = (args ?? {}) as { slideIndex: number; shapeIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        const slide = slides.items[slideIndex];
        slide.shapes.load('items');
        await context.sync();
        const shapeCount = slide.shapes.items.length;
        if (shapeIndex < 0 || shapeIndex >= shapeCount) {
          return `Invalid shapeIndex ${String(shapeIndex)}. Slide has ${String(shapeCount)} shapes.`;
        }
        slide.shapes.items[shapeIndex].delete();
        await context.sync();
        return `Deleted shape ${String(shapeIndex)} from slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSlideLayouts: Tool = {
  name: 'get_slide_layouts',
  description: 'List all available slide layouts from the first slide master. Use with apply_slide_layout.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await PowerPoint.run(async context => {
        const masters = context.presentation.slideMasters;
        masters.load('items');
        await context.sync();
        if (masters.items.length === 0) return 'No slide masters found.';

        const master = masters.items[0];
        master.layouts.load('items');
        await context.sync();

        const lines = [`Slide Layouts (${String(master.layouts.items.length)})`, '='.repeat(40)];
        for (let i = 0; i < master.layouts.items.length; i++) {
          const layout = master.layouts.items[i];
          layout.load('name,id');
          await context.sync();
          lines.push(`${String(i)}. "${layout.name}" (id: ${layout.id})`);
        }
        return lines.join('\n');
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const applySlideLayout: Tool = {
  name: 'apply_slide_layout',
  description: 'Apply a layout from the first slide master to a slide. Use get_slide_layouts to see available layouts.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      layoutIndex: { type: 'number', description: '0-based layout index from get_slide_layouts.' },
    },
    required: ['slideIndex', 'layoutIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, layoutIndex } = (args ?? {}) as { slideIndex: number; layoutIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const masters = context.presentation.slideMasters;
        masters.load('items');
        await context.sync();
        if (masters.items.length === 0) return 'No slide masters found.';

        const master = masters.items[0];
        master.layouts.load('items');
        await context.sync();
        if (layoutIndex < 0 || layoutIndex >= master.layouts.items.length) {
          return `Invalid layoutIndex ${String(layoutIndex)}. Must be 0-${String(master.layouts.items.length - 1)}.`;
        }

        const layout = master.layouts.items[layoutIndex];
        layout.load('name');
        await context.sync();

        slides.items[slideIndex].applyLayout(layout);
        await context.sync();
        return `Applied layout "${layout.name}" to slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelectedSlides: Tool = {
  name: 'get_selected_slides',
  description: 'Get the currently selected slides. Returns both the 1-based slide number (what the user sees) and the 0-based index (for use with other tools).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await PowerPoint.run(async context => {
        const selected = context.presentation.getSelectedSlides();
        selected.load('items');
        await context.sync();
        if (selected.items.length === 0) return 'No slides selected.';
        const info: string[] = [];
        for (const slide of selected.items) {
          slide.load('index');
          await context.sync();
          // slide.index is 1-based; tools use 0-based slideIndex
          info.push(`Slide ${String(slide.index)} (use slideIndex=${String(slide.index - 1)} in tools)`);
        }
        return `Currently on: ${info.join(', ')}`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelectedShapes: Tool = {
  name: 'get_selected_shapes',
  description: 'Get the currently selected shapes on the active slide.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await PowerPoint.run(async context => {
        const selected = context.presentation.getSelectedShapes();
        selected.load('items');
        await context.sync();
        if (selected.items.length === 0) return 'No shapes selected.';
        const lines = [`Selected shapes (${String(selected.items.length)}):`];
        for (const shape of selected.items) {
          shape.load('name,left,top,width,height');
          await context.sync();
          lines.push(`- "${shape.name}" at (${String(Math.round(shape.left))}, ${String(Math.round(shape.top))}) size ${String(Math.round(shape.width))}×${String(Math.round(shape.height))}`);
        }
        return lines.join('\n');
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setShapeText: Tool = {
  name: 'set_shape_text',
  description: 'Set the text of an existing shape by index, preserving the shape itself. Unlike update_slide_shape, this tool also supports adding text to shapes that currently have no text.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index.' },
      text: { type: 'string', description: 'The text to set.' },
    },
    required: ['slideIndex', 'shapeIndex', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, shapeIndex, text } = (args ?? {}) as {
      slideIndex: number;
      shapeIndex: number;
      text: string;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }
        const slide = slides.items[slideIndex];
        slide.shapes.load('items');
        await context.sync();
        const shapeCount = slide.shapes.items.length;
        if (shapeIndex < 0 || shapeIndex >= shapeCount) {
          return `Invalid shapeIndex ${String(shapeIndex)}. Slide has ${String(shapeCount)} shapes.`;
        }
        const shape = slide.shapes.items[shapeIndex];
        shape.textFrame.textRange.text = text;
        await context.sync();
        return `Set text of shape ${String(shapeIndex)} on slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

export const powerPointTools: Tool[] = [
  getPresentationOverview,
  getPresentationContent,
  getSlideImage,
  getSlideNotes,
  getSlideShapes,
  getSlideLayouts,
  getSelectedSlides,
  getSelectedShapes,
  setPresentationContent,
  addSlideFromCode,
  addGeometricShape,
  addLine,
  clearSlide,
  deleteSlide,
  moveSlide,
  setSlideBackground,
  applySlideLayout,
  updateSlideShape,
  setShapeText,
  updateShapeStyle,
  moveResizeShape,
  deleteShape,
  setSlideNotes,
  duplicateSlide,
];
