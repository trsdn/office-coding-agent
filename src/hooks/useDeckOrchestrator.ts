/**
 * Deck Orchestrator
 *
 * Manages a planner → worker multi-session flow for creating slide decks.
 * Each worker gets a fresh context window to avoid context exhaustion.
 */

import type { WebSocketCopilotClient } from '@/lib/websocket-client';
import type { SessionEvent } from '@github/copilot-sdk';
import { runSubSession } from '@/lib/session-factory';
import { submitPlanTool, extractPlanFromEvents } from '@/tools/planner';
import type { DeckPlan, SlidePlan } from '@/tools/planner';
import { getToolsForHost } from '@/tools';
import plannerPromptRaw from '@/services/ai/prompts/PLANNER_PROMPT.md?raw';
import workerPromptRaw from '@/services/ai/prompts/WORKER_PROMPT.md?raw';

export type SlideStatus = 'pending' | 'running' | 'done' | 'failed';

export interface SlideProgress {
  plan: SlidePlan;
  status: SlideStatus;
  error?: string;
}

export interface DeckOrchestratorCallbacks {
  /** Called when the planner finishes with a plan */
  onPlan: (plan: DeckPlan) => void;
  /** Called when a slide's status changes */
  onSlideProgress: (index: number, progress: SlideProgress) => void;
  /** Called with streaming text from planner or workers */
  onText: (text: string) => void;
  /** Called with tool events from workers (for UI progress indicators) */
  onWorkerEvent?: (slideIndex: number, event: SessionEvent) => void;
  /** Called when the entire deck is complete */
  onComplete: (results: SlideProgress[]) => void;
  /** Called on fatal error */
  onError: (error: string) => void;
}

/**
 * Build the worker prompt for a specific slide task.
 */
function buildWorkerPrompt(slide: SlidePlan, totalSlides: number): string {
  return `${workerPromptRaw}

## Slide Task

Create slide ${String(slide.index + 1)} of ${String(totalSlides)}:
- **Title**: ${slide.title}
- **Layout**: ${slide.layout}
- **Content**: ${slide.content}

Narrate your progress in the user's language. Start with "Slide ${String(slide.index + 1)}/${String(totalSlides)}: ${slide.title}…"`;
}

/**
 * Run the full planner → worker orchestration.
 *
 * @param client - Shared WebSocket Copilot client
 * @param model - Model ID to use for all sessions
 * @param userPrompt - The user's original deck request
 * @param callbacks - Progress callbacks for UI updates
 * @param signal - Optional AbortSignal for cancellation
 */
export async function orchestrateDeck(
  client: WebSocketCopilotClient,
  model: string,
  userPrompt: string,
  callbacks: DeckOrchestratorCallbacks,
  signal?: AbortSignal,
): Promise<void> {
  // --- Phase 1: Planner ---
  callbacks.onText('📋 Erstelle Plan…\n');

  const plannerResult = await runSubSession(
    client,
    {
      model,
      systemPrompt: plannerPromptRaw,
      tools: [submitPlanTool],
    },
    userPrompt,
    (event) => {
      if (event.type === 'assistant.message_delta') {
        callbacks.onText(event.data.deltaContent);
      }
    },
  );

  if (!plannerResult.success) {
    callbacks.onError(`Planner failed: ${plannerResult.error ?? 'unknown error'}`);
    return;
  }

  // Extract plan from tool call events
  const plan = extractPlanFromEvents(
    plannerResult.events as Array<{ type: string; data: Record<string, unknown> }>,
  );

  if (!plan || plan.slides.length === 0) {
    callbacks.onError('Planner did not produce a valid slide plan.');
    return;
  }

  callbacks.onPlan(plan);
  callbacks.onText(`\n\n📋 Plan: ${String(plan.slides.length)} Slides\n`);

  // --- Phase 2: Workers ---
  const results: SlideProgress[] = plan.slides.map(slide => ({
    plan: slide,
    status: 'pending' as SlideStatus,
  }));

  const pptTools = getToolsForHost('powerpoint');

  for (let i = 0; i < plan.slides.length; i++) {
    if (signal?.aborted) {
      callbacks.onText('\n⚠️ Abgebrochen.\n');
      break;
    }

    const slide = plan.slides[i];
    results[i].status = 'running';
    callbacks.onSlideProgress(i, results[i]);
    callbacks.onText(`\n🔄 Slide ${String(i + 1)}/${String(plan.slides.length)}: ${slide.title}\n`);

    const workerPrompt = buildWorkerPrompt(slide, plan.slides.length);

    const workerResult = await runSubSession(
      client,
      {
        model,
        systemPrompt: workerPrompt,
        tools: pptTools,
      },
      `Create slide ${String(slide.index)}: "${slide.title}" — Layout: ${slide.layout}. Content: ${slide.content}`,
      (event) => {
        callbacks.onWorkerEvent?.(i, event);
        if (event.type === 'assistant.message_delta') {
          callbacks.onText(event.data.deltaContent);
        }
      },
    );

    if (workerResult.success) {
      results[i].status = 'done';
      callbacks.onText(`\n✅ Slide ${String(i + 1)} fertig\n`);
    } else {
      // Retry once
      callbacks.onText(`\n⚠️ Slide ${String(i + 1)} fehlgeschlagen, versuche erneut…\n`);
      const retryResult = await runSubSession(
        client,
        { model, systemPrompt: workerPrompt, tools: pptTools },
        `Create slide ${String(slide.index)}: "${slide.title}" — Layout: ${slide.layout}. Content: ${slide.content}`,
        (event) => {
          callbacks.onWorkerEvent?.(i, event);
          if (event.type === 'assistant.message_delta') {
            callbacks.onText(event.data.deltaContent);
          }
        },
      );

      if (retryResult.success) {
        results[i].status = 'done';
        callbacks.onText(`\n✅ Slide ${String(i + 1)} fertig (2. Versuch)\n`);
      } else {
        results[i].status = 'failed';
        results[i].error = retryResult.error;
        callbacks.onText(`\n❌ Slide ${String(i + 1)} fehlgeschlagen: ${retryResult.error ?? 'unknown'}\n`);
      }
    }

    callbacks.onSlideProgress(i, results[i]);
  }

  // --- Phase 3: Summary ---
  const done = results.filter(r => r.status === 'done').length;
  const failed = results.filter(r => r.status === 'failed').length;
  callbacks.onText(`\n\n📊 Fertig: ${String(done)}/${String(plan.slides.length)} Slides erstellt`);
  if (failed > 0) {
    callbacks.onText(` (${String(failed)} fehlgeschlagen)`);
  }
  callbacks.onText('\n');

  callbacks.onComplete(results);
}
