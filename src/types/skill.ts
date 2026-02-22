/** Parsed metadata from a skill's YAML frontmatter. */
export interface SkillMetadata {
  name: string;
  description: string;
  version: string;
  tags: string[];
  /** Office hosts where this skill is relevant. If empty/omitted, shown for all hosts. */
  hosts?: string[];
  license?: string;
  repository?: string;
  documentation?: string;
}

/** A loaded agent skill with metadata and content. */
export interface AgentSkill {
  /** Parsed YAML frontmatter */
  metadata: SkillMetadata;
  /** Markdown body (without frontmatter) — injected into system prompt */
  content: string;
}
