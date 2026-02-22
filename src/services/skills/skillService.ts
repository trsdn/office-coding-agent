import type { AgentSkill, SkillMetadata } from '@/types/skill';
import excelSkillRaw from '@/skills/excel/SKILL.md';
import outlookSkillRaw from '@/skills/outlook/SKILL.md';
import powerpointSkillRaw from '@/skills/powerpoint/SKILL.md';
import powerpointDeckBuilderSkillRaw from '@/skills/powerpoint-deck-builder/SKILL.md';
import powerpointRedesignSkillRaw from '@/skills/powerpoint-redesign/SKILL.md';
import powerpointFormattingSkillRaw from '@/skills/powerpoint-formatting/SKILL.md';
import wordSkillRaw from '@/skills/word/SKILL.md';

/**
 * Parse YAML frontmatter from a skill markdown file.
 * Handles the `---` delimited block at the top of the file.
 */
export function parseFrontmatter(raw: string): { metadata: SkillMetadata; content: string } {
  const trimmed = raw.trimStart();

  if (!trimmed.startsWith('---')) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [] },
      content: trimmed,
    };
  }

  const endIndex = trimmed.indexOf('---', 3);
  if (endIndex === -1) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [] },
      content: trimmed,
    };
  }

  const yamlBlock = trimmed.slice(3, endIndex).trim();
  const content = trimmed.slice(endIndex + 3).trim();

  // Simple YAML parser for flat key-value and tag arrays
  const metadata: SkillMetadata = {
    name: '',
    description: '',
    version: '0.0.0',
    tags: [],
  };

  let currentKey = '';
  let isMultilineValue = false;
  let multilineValue = '';

  for (const line of yamlBlock.split('\n')) {
    const trimmedLine = line.trim();

    // Array items: "  - value"
    if (trimmedLine.startsWith('- ') && currentKey === 'tags') {
      const itemValue = trimmedLine.slice(2).trim();
      metadata.tags.push(itemValue);
      continue;
    }

    // Multiline continuation (indented lines after "key: >")
    if (isMultilineValue && (line.startsWith('  ') || line.startsWith('\t'))) {
      multilineValue += (multilineValue ? ' ' : '') + trimmedLine;
      continue;
    }

    // Flush multiline value
    if (isMultilineValue) {
      setMetadataField(metadata, currentKey, multilineValue);
      isMultilineValue = false;
      multilineValue = '';
    }

    // Key-value pairs
    const colonIndex = trimmedLine.indexOf(':');
    if (colonIndex === -1) continue;

    currentKey = trimmedLine.slice(0, colonIndex).trim();
    const value = trimmedLine.slice(colonIndex + 1).trim();

    if (value === '>' || value === '|') {
      // Multiline scalar
      isMultilineValue = true;
      multilineValue = '';
    } else if (value === '') {
      // Could be start of array (tags:) — handled by "- " check above
      continue;
    } else {
      setMetadataField(metadata, currentKey, value);
    }
  }

  // Flush any trailing multiline value
  if (isMultilineValue && multilineValue) {
    setMetadataField(metadata, currentKey, multilineValue);
  }

  return { metadata, content };
}

function setMetadataField(metadata: SkillMetadata, key: string, value: string): void {
  switch (key) {
    case 'name':
      metadata.name = value;
      break;
    case 'description':
      metadata.description = value;
      break;
    case 'version':
      metadata.version = value;
      break;
    case 'license':
      metadata.license = value;
      break;
    case 'repository':
      metadata.repository = value;
      break;
    case 'documentation':
      metadata.documentation = value;
      break;
    case 'hosts':
      metadata.hosts = value
        .replace(/^\[/, '')
        .replace(/\]$/, '')
        .split(',')
        .map(h => h.trim())
        .filter(Boolean);
      break;
  }
}

function loadBundledSkills(): AgentSkill[] {
  const bundledRawSkills = [excelSkillRaw, outlookSkillRaw, powerpointSkillRaw, powerpointDeckBuilderSkillRaw, powerpointRedesignSkillRaw, powerpointFormattingSkillRaw, wordSkillRaw];

  const loaded = bundledRawSkills.map(raw => {
    const parsed = parseFrontmatter(raw);
    return { metadata: parsed.metadata, content: parsed.content };
  });

  return loaded.sort((left, right) => left.metadata.name.localeCompare(right.metadata.name));
}

const bundledSkills: AgentSkill[] = loadBundledSkills();
let importedSkills: AgentSkill[] = [];

export function getBundledSkills(): AgentSkill[] {
  return bundledSkills;
}

export function getImportedSkills(): AgentSkill[] {
  return importedSkills;
}

export function setImportedSkills(skills: AgentSkill[]): void {
  importedSkills = skills;
}

/**
 * Get all loaded agent skills, optionally filtered by host.
 * Skills with no `hosts` field are shown for all hosts.
 */
export function getSkills(host?: string): AgentSkill[] {
  const all = [...bundledSkills, ...importedSkills];
  if (!host) return all;
  return all.filter(
    s => !s.metadata.hosts || s.metadata.hosts.length === 0 || s.metadata.hosts.includes(host)
  );
}

/**
 * Get a single skill by name.
 */
export function getSkill(name: string): AgentSkill | undefined {
  return getSkills().find(s => s.metadata.name === name);
}

/**
 * Build the combined skill context string for injection into the system prompt.
 * @param activeNames — if provided, only include skills whose names are in this list.
 *                       If omitted or empty, all bundled skills are included.
 * @param host — if provided, only include skills that match this host.
 * Returns an empty string if no skills match.
 */
export function buildSkillContext(activeNames?: string[], host?: string): string {
  let skills = getSkills(host);

  // Filter to active names if provided (empty array = none active)
  if (activeNames !== undefined) {
    skills = skills.filter(s => activeNames.includes(s.metadata.name));
  }

  if (skills.length === 0) return '';

  const sections = skills.map(
    skill =>
      `\n\n---\n## Agent Skill: ${skill.metadata.name}\n${skill.metadata.description}\n\n${skill.content}`
  );

  return `\n\n# Agent Skills\nThe following agent skills provide domain-specific knowledge. Use them to help the user with specialized tasks.${sections.join('')}`;
}
