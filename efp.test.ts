import { expect, test, describe } from "bun:test";
import {
  ExcelFormulaParser,
  TokenType,
  TokenSubType,
  type FormulaToken,
} from "./efp";

describe("ExcelFormulaParser", () => {
  const testFormulae = [
    { formula: "=SUM())", description: "Unbalanced parentheses" },
    { formula: '=SUM("")', description: "Empty string argument" },
    { formula: '="あいうえお"&H3&"b"', description: "Unicode characters" },
    { formula: "=1+3+5", description: "Simple addition" },
    { formula: "=3 * 4 + 5", description: "Mixed operators" },
    { formula: "=50", description: "Single number" },
    { formula: "=1+1", description: "Simple addition" },
    { formula: "=$A1", description: "Relative column reference" },
    { formula: "=$B$2", description: "Absolute reference" },
    { formula: "=SUM(B5:B15)", description: "Simple range" },
    { formula: "=SUM(B5:B15,D5:D15)", description: "Multiple ranges" },
    { formula: "=SUM(sheet1!$A$1:$B$2)", description: "Sheet reference" },
    {
      formula: "=[data.xls]sheet1!$A$1",
      description: "External workbook reference",
    },
    { formula: "=#]#NUM!", description: "Error value" },
    { formula: "=3.1E-24-2.1E-24", description: "Scientific notation" },
    { formula: "=1%2", description: "Percentage operator" },
    { formula: "={1,2}", description: "Array literal" },
    { formula: "=TRUE", description: "Boolean value" },
    { formula: "=--1-1", description: "Multiple unary operators" },
    { formula: "=10*2^(2*(1+1))%", description: "Complex operator precedence" },
  ];

  describe("Formula Parsing", () => {
    for (const { formula, description } of testFormulae) {
      test(`should parse ${description} (${formula})`, () => {
        const parser = new ExcelFormulaParser();
        const tokens = parser.parse(formula.replace(/^=/, ""));

        // Basic structure validation
        expect(tokens.length).toBeGreaterThan(0);
        expect(tokens.every((t) => typeof t.value === "string")).toBe(true);

        // Validate start/stop pairs
        const stack: FormulaToken[] = [];
        for (const token of tokens) {
          if (token.subtype === TokenSubType.Start) stack.push(token);
          if (token.subtype === TokenSubType.Stop) stack.pop();
        }
        expect(stack.length).toBe(0);
      });
    }
  });

  describe("Pretty Print", () => {
    for (const { formula, description } of testFormulae) {
      test(`should generate pretty output for ${description}`, () => {
        const parser = new ExcelFormulaParser();
        const tokens = parser.parse(formula.replace(/^=/, ""));
        const output = parser.prettyPrint(tokens);

        // Basic output validation
        expect(output).toContain("<");
        expect(output).toContain(">");
        expect(output.split("\n").length).toBeGreaterThan(1);
      });
    }
  });

  describe("Formula Rendering", () => {
    for (const { formula, description } of testFormulae) {
      test(`should re-render ${description} correctly`, () => {
        const parser = new ExcelFormulaParser();
        const normalizedFormula = formula.replace(/^=/, "").replace(/\s+/g, "");
        const tokens = parser.parse(normalizedFormula);
        const rendered = parser.render(tokens);

        // Normalize both for comparison
        const normalizedRendered = rendered.replace(/\s+/g, "");
        expect(normalizedRendered).toBe(normalizedFormula);
      });
    }
  });

  describe("Edge Cases", () => {
    const edgeCases = [
      { input: "", expected: 0 },
      { input: "+", expected: 0 },
      { input: "                 ", expected: 0 },
      { input: "=IF(R#", expected: 3 }, // Partial formula
      { input: "=IF(R{", expected: 3 }, // Partial array
      { input: `=""+'''`, expected: 3 }, // Mixed quotes
    ];

    for (const { input, expected } of edgeCases) {
      test(`should handle "${input}"`, () => {
        const parser = new ExcelFormulaParser();
        const tokens = parser.parse(input.replace(/^=/, ""));

        if (expected === 0) {
          expect(tokens.length).toBe(0);
        } else {
          expect(tokens.length).toBe(expected);
        }
      });
    }
  });

  describe("Complex Formulae", () => {
    const complexFormulae = [
      `=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))`,
      "={SUM(B2:D2*B3:D3)}",
      "=SUM(123 + SUM(456) + (45<6))+456+789",
      `=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") & " more text"`,
    ];

    for (const formula of complexFormulae) {
      test(`should parse complex formula: ${formula.substring(
        0,
        20
      )}...`, () => {
        const parser = new ExcelFormulaParser();
        const tokens = parser.parse(formula.replace(/^=/, ""));

        // Verify function nesting
        const functionStack: FormulaToken[] = [];
        for (const token of tokens) {
          if (
            token.type === TokenType.Function &&
            token.subtype === TokenSubType.Start
          ) {
            functionStack.push(token);
          }
          if (
            token.type === TokenType.Function &&
            token.subtype === TokenSubType.Stop
          ) {
            functionStack.pop();
          }
        }
        expect(functionStack.length).toBe(0);
        expect(tokens.some((t) => t.type === TokenType.OperatorInfix)).toBe(
          true
        );
      });
    }
  });
});
