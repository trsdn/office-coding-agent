import React, { useState } from 'react';
import { RotateCcw, ServerIcon, BrainCircuit } from 'lucide-react';
import { SkillPicker } from './SkillPicker';
import { McpManagerDialog } from './McpManagerDialog';
import { useSettingsStore } from '@/stores';

export interface ChatHeaderProps {
  onClearMessages: () => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({ onClearMessages }) => {
  const [mcpOpen, setMcpOpen] = useState(false);
  const workiqEnabled = useSettingsStore(s => s.workiqEnabled);
  const toggleWorkiq = useSettingsStore(s => s.toggleWorkiq);

  return (
    <div className="flex items-center justify-between border-b border-border bg-background px-3 py-1.5">
      <div className="flex items-center gap-2 min-w-0">
        <SkillPicker />
      </div>

      <div className="flex items-center gap-0.5">
        <button
          onClick={toggleWorkiq}
          className={`inline-flex h-8 items-center gap-1 rounded-md px-2 text-xs font-medium transition-colors ${
            workiqEnabled
              ? 'bg-primary/10 text-primary hover:bg-primary/20'
              : 'text-muted-foreground hover:bg-accent hover:text-accent-foreground'
          }`}
          aria-label={`WorkIQ: ${workiqEnabled ? 'on' : 'off'}`}
          aria-pressed={workiqEnabled}
          title={
            workiqEnabled
              ? 'WorkIQ enabled — click to disable'
              : 'Enable WorkIQ (Microsoft 365 data)'
          }
        >
          <BrainCircuit className="size-4" />
          <span>WorkIQ</span>
        </button>
        <button
          onClick={() => setMcpOpen(true)}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="MCP Servers"
          title="MCP Servers"
        >
          <ServerIcon className="size-4" />
        </button>
        <button
          onClick={onClearMessages}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="New conversation"
          title="New conversation"
        >
          <RotateCcw className="size-4" />
        </button>
      </div>

      <McpManagerDialog open={mcpOpen} onOpenChange={setMcpOpen} />
    </div>
  );
};
