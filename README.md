# Excel Formula Parser

A TypeScript library for parsing and tokenizing Excel formulas. This parser handles complex Excel formulas, including nested functions, arrays, and various operators, making it ideal for applications that need to analyze or manipulate Excel formulas programmatically.

## Features

- 📊 Parses complex Excel formulas into tokens
- 🔄 Supports formula reconstruction from tokens
- 🎯 Handles various Excel-specific elements:
  - Function calls and nested expressions
  - Array literals and range references
  - Sheet references and external workbook references
  - Error values (#NULL!, #DIV/0!, etc.)
  - Text strings with Unicode support
  - Scientific notation
  - Boolean values
  - Operators (mathematical, logical, concatenation)

## Installation

```bash
npm install excel-formula-parser
# or
yarn add excel-formula-parser
# or
pnpm add excel-formula-parser
or
bun install excel-formula-parser
```

## Usage

### Basic Usage

```typescript
import { ExcelFormulaParser } from 'excel-formula-parser';

const parser = new ExcelFormulaParser();

// Parse a formula (the '=' is optional)
const tokens = parser.parse('=SUM(A1:A10)');

// Pretty print the tokens for debugging
console.log(parser.prettyPrint(tokens));

// Render tokens back to a formula string
const formula = parser.render(tokens);
```

### Token Structure

Each token has the following structure:

```typescript
interface FormulaToken {
  value: string;    // The actual value of the token
  type: TokenType;  // The type of token (e.g., Function, Operand, Operator)
  subtype: TokenSubType | ""; // Additional type information
}
```

### Examples

Here are some examples of formulas the parser can handle:

```typescript
const parser = new ExcelFormulaParser();

// Simple arithmetic
parser.parse('1+2*3');

// Function calls
parser.parse('SUM(B5:B15)');

// Nested functions
parser.parse('IF(P5=1.0,"NA",IF(P5=2.0,"A","B"))');

// Array literals
parser.parse('{1,2;3,4}');

// External references
parser.parse('[data.xls]sheet1!$A$1');

// Complex formulas
parser.parse('IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") & " more text"');
```

### Pretty Printing

The `prettyPrint` method helps visualize the token structure:

```typescript
const parser = new ExcelFormulaParser();
const tokens = parser.parse('SUM(A1:A10,B1)');
console.log(parser.prettyPrint(tokens));

// Output:
// SUM <Function> <Start>
//   A1:A10 <Operand> <Range>
//   , <Argument> <>
//   B1 <Operand> <Range>
// ) <Function> <Stop>
```

### Token Types

The parser recognizes various token types:

```typescript
enum TokenType {
  Noop = "Noop",
  Operand = "Operand",
  Function = "Function",
  Subexpression = "Subexpression",
  Argument = "Argument",
  OperatorPrefix = "OperatorPrefix",
  OperatorInfix = "OperatorInfix",
  OperatorPostfix = "OperatorPostfix",
  Whitespace = "Whitespace",
  Unknown = "Unknown"
}
```

And subtypes:

```typescript
enum TokenSubType {
  Start = "Start",
  Stop = "Stop",
  Text = "Text",
  Number = "Number",
  Logical = "Logical",
  Error = "Error",
  Range = "Range",
  Math = "Math",
  Concatenation = "Concatenation",
  Intersection = "Intersection",
  Union = "Union"
}
```

## Special Cases Handled

- Unicode characters in strings
- Scientific notation
- External workbook references
- Sheet names with special characters
- Error values
- Array formulas
- Complex nested functions
- Multiple ranges and union operators
- Absolute and relative cell references

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## Testing

The library comes with a comprehensive test suite built with Bun. To run the tests:

```bash
bun test
```

## License

[MIT](LICENSE)

## Acknowledgments

This parser is inspired mainly by [efp](https://github.com/xuri/efp) in Go and aims to provide a robust, TypeScript-first approach to formula parsing.