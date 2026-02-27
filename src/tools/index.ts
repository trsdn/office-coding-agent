import { createTools } from './codegen';
import {
  rangeConfigs,
  rangeFormatConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
} from './configs';
import type { ToolConfig, ToolConfigBase } from './codegen/types';
import type { Tool } from '@github/copilot-sdk';
import type { OfficeHostApp } from '@/services/office/host';
import { powerPointTools } from './powerpoint';
import { wordTools } from './word';
import { outlookTools } from './outlook';
import { managementTools } from './management';

export { managementTools } from './management';

export const MAX_TOOLS_PER_REQUEST = 128;

/** All Excel tool configs combined for manifest generation */
export const allConfigs: readonly (readonly ToolConfig[])[] = [
  rangeConfigs,
  rangeFormatConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
];

/** All tool configs across all hosts — for manifest generation */
export const allConfigsByHost: Record<string, readonly (readonly ToolConfigBase[])[]> = {
  excel: allConfigs,
};

/** All Excel tools combined into a single array for Copilot SDK */
export const excelTools: Tool[] = allConfigs.flatMap(configs => createTools(configs));

export { powerPointTools } from './powerpoint';
export { wordTools } from './word';
export { outlookTools } from './outlook';

/** General-purpose tools included for all hosts */
const generalTools: Tool[] = [...managementTools];

export function getToolsForHost(host: OfficeHostApp): Tool[] {
  let hostTools: Tool[];
  switch (host) {
    case 'excel':
      hostTools = excelTools;
      break;
    case 'powerpoint':
      hostTools = powerPointTools;
      break;
    case 'word':
      hostTools = wordTools;
      break;
    case 'outlook':
      hostTools = outlookTools;
      break;
    default:
      return [];
  }
  return [...hostTools, ...generalTools].slice(0, MAX_TOOLS_PER_REQUEST);
}
