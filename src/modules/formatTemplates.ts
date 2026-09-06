/**
 * Format Template Module
 *
 * Provides manual input format templates that users can configure.
 * The user inputs values that are assembled into a formatted output string.
 *
 * Pattern 1 (from image row 3):
 *   "Q [Q17] Q17. SPONTANEOUS LIKES"
 *   Template: "Q [" + input1 + "] " + input2
 *   Inputs: ["Q17", "Q17. SPONTANEOUS LIKES"]
 *
 * Pattern 2 (from image row 5):
 *   "L 1066L3R999"
 *   Template: "L" + input1 + "L" + input2 + "R" + input3
 *   Inputs: ["1066", "3", "999"]
 */

export interface FormatTemplate {
  id: string;
  name: string;
  description: string;
  // The output format as a string with placeholders
  // e.g. "Q [{{1}}] {{2}}" or "L{{1}}L{{2}}R{{3}}"
  templateString: string;
  // How many input fields the user needs
  inputCount: number;
  // Example input values
  exampleInputs: string[];
  // Human-readable description of how inputs combine
  formatDescription: string;
}

export const FORMAT_TEMPLATES: FormatTemplate[] = [
  {
    id: 'q-pattern',
    name: 'Q-Pattern (Row 3 style)',
    description: 'Used for question/text rows like: Q [Q17] Q17. SPONTANEOUS LIKES',
    templateString: 'Q [{{1}}] {{2}}',
    inputCount: 2,
    exampleInputs: ['Q17', 'Q17. SPONTANEOUS LIKES'],
    formatDescription: 'Inputs: "Q17" + "Q17. SPONTANEOUS LIKES" → "Q [Q17] Q17. SPONTANEOUS LIKES"',
  },
  {
    id: 'l-pattern',
    name: 'L-Pattern (Row 5 style)',
    description: 'Used for numeric ID rows like: L 1L3R999',
    templateString: 'L {{1}}L{{2}}R{{3}}',
    inputCount: 3,
    exampleInputs: ['1', '3', '999'],
    formatDescription: 'Inputs: "1" + "3" + "999" → "L 1L3R999"',
  },
];

/**
 * A user-defined format instance with their input values.
 */
export interface UserFormatInstance {
  templateId: string;
  inputValues: string[]; // User-entered values
  // The column this format applies to (user selects which column uses this format)
  targetColumnKey: string;
}

/**
 * Build the formatted output string from a template and user inputs.
 */
export function buildFormattedString(
  template: FormatTemplate,
  inputs: string[]
): string {
  let result = template.templateString;
  for (let i = 0; i < Math.min(inputs.length, template.inputCount); i++) {
    const placeholder = `{{${i + 1}}}`;
    const value = inputs[i] || '';
    result = result.replace(new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), value);
  }
  return result;
}

/**
 * Get placeholder label for an input field.
 */
export function getInputLabel(template: FormatTemplate, index: number): string {
  switch (template.id) {
    case 'q-pattern':
      return index === 0 ? 'Q-Number (e.g. Q17)' : 'Question Text (e.g. Q17. SPONTANEOUS LIKES)';
    case 'l-pattern':
      return index === 0 ? 'Code 1 (e.g. 1066)' : index === 1 ? 'Code 2 (e.g. 3)' : 'Code 3 (e.g. 999)';
    default:
      return `Input ${index + 1}`;
  }
}
