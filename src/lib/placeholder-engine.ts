import { formatCellValue } from './excel-parser';

export interface SubstitutionContext {
  certificateId: string;
  studentId: string;
  issueDate: string;
  eventDate?: string;
  eventName: string;
  row: Record<string, any>;
}

export function substitutePlaceholders(
  templateHtml: string,
  context: SubstitutionContext
): string {
  if (!templateHtml) return '';

  let result = templateHtml;

  // 1. Replace reserved system tokens first
  const systemMap: Record<string, string> = {
    'certificate_id': context.certificateId,
    'certificate_code': context.certificateId,
    'student_id': context.studentId,
    'issue_date': context.issueDate,
    'event_date': context.eventDate || context.issueDate,
    'event_name': context.eventName,
  };

  Object.entries(systemMap).forEach(([key, val]) => {
    const regex = new RegExp(`<<\\s*${key}\\s*>>`, 'gi');
    result = result.replace(regex, val);
  });

  // 2. Replace dynamic Excel column headers
  result = result.replace(/<<\s*([^>]+?)\s*>>/g, (match, rawTokenName) => {
    const tokenName = rawTokenName.trim();
    
    // Check if matching system token again
    const sysKey = tokenName.toLowerCase();
    if (systemMap[sysKey] !== undefined) {
      return systemMap[sysKey];
    }

    // Try finding in row by exact raw key or trimmed key
    for (const [rowKey, rowVal] of Object.entries(context.row)) {
      if (rowKey === tokenName || rowKey.trim() === tokenName) {
        return formatCellValue(rowVal);
      }
    }

    // Case-insensitive match fallback
    for (const [rowKey, rowVal] of Object.entries(context.row)) {
      if (rowKey.trim().toLowerCase() === tokenName.toLowerCase()) {
        return formatCellValue(rowVal);
      }
    }

    return match;
  });

  return result;
}

/**
 * Scans template HTML and extracts all << token >> names used
 */
export function extractTemplateTokens(templateHtml: string): string[] {
  const matches = templateHtml.match(/<<\s*([^>]+?)\s*>>/g) || [];
  const tokens = new Set<string>();
  matches.forEach(m => {
    const inner = m.replace(/^<<\s*/, '').replace(/\s*>>$/, '').trim();
    if (inner) tokens.add(inner);
  });
  return Array.from(tokens);
}

/**
 * Validates which template tokens are missing in the current Excel dataset
 */
export function validateTemplateTokens(
  templateHtml: string,
  availableHeaders: string[]
): {
  allTokens: string[];
  matchedTokens: string[];
  unmatchedTokens: string[];
  systemTokens: string[];
} {
  const allTokens = extractTemplateTokens(templateHtml);
  const systemReserved = ['certificate_id', 'certificate_code', 'student_id', 'issue_date', 'event_date', 'event_name'];
  
  const systemTokens: string[] = [];
  const matchedTokens: string[] = [];
  const unmatchedTokens: string[] = [];

  const normalizedAvailable = availableHeaders.map(h => h.trim().toLowerCase());

  allTokens.forEach(token => {
    const tokenLower = token.toLowerCase();
    if (systemReserved.includes(tokenLower)) {
      systemTokens.push(token);
    } else if (normalizedAvailable.includes(tokenLower)) {
      matchedTokens.push(token);
    } else {
      unmatchedTokens.push(token);
    }
  });

  return {
    allTokens,
    matchedTokens,
    unmatchedTokens,
    systemTokens
  };
}
