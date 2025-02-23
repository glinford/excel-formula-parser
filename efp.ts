// Token types and subtypes
export enum TokenType {
	Noop = "Noop",
	Operand = "Operand",
	Function = "Function",
	Subexpression = "Subexpression",
	Argument = "Argument",
	OperatorPrefix = "OperatorPrefix",
	OperatorInfix = "OperatorInfix",
	OperatorPostfix = "OperatorPostfix",
	Whitespace = "Whitespace",
	Unknown = "Unknown",
}

export enum TokenSubType {
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
	Union = "Union",
}

export interface FormulaToken {
	value: string;
	type: TokenType;
	subtype: TokenSubType | "";
}

export class ExcelFormulaParser {
	private formula: string = "";
	private chars: string[] = [];
	private tokens: FormulaToken[] = [];
	private stack: FormulaToken[] = [];
	private offset: number = 0;
	private currentToken: string[] = [];

	// State flags
	private inString: boolean = false;
	private inPath: boolean = false;
	private inRange: boolean = false;
	private inError: boolean = false;

	public parse(formula: string): FormulaToken[] {
		if (!formula.trim() || ["+", "-", "*", "/"].includes(formula))
			return [];
		this.resetState();
		this.formula = formula.replace(/^=/, "");
		this.chars = [...this.formula];

		while (this.offset < this.chars.length) {
			this.processNextCharacter();
		}

		this.finalizeCurrentToken();
		return this.postProcessTokens();
	}

	private finalizeCurrentToken(): void {
		if (this.currentToken.length > 0) {
			const value = this.currentToken.join("");
			this.currentToken = [];

			if (this.isNumber(value)) {
				this.addToken(TokenType.Operand, TokenSubType.Number, value);
			} else if (this.isLogical(value)) {
				this.addToken(TokenType.Operand, TokenSubType.Logical, value);
			} else {
				this.addToken(TokenType.Operand, TokenSubType.Range, value);
			}
		}
	}

	private handleOpenBrace(): void {
		this.finalizeCurrentToken();
		this.stack.push(
			this.addToken(TokenType.Function, TokenSubType.Start, "ARRAY"),
		);
		this.offset++;
	}

	private handleCloseBrace(): void {
		this.finalizeCurrentToken();
		this.addToken(TokenType.Function, TokenSubType.Stop, "ARRAY");
		if (this.stack.length > 0) {
			this.stack.pop();
		}
		this.offset++;
	}

	private handleHash(): void {
		this.finalizeCurrentToken();
		this.inError = true;
		this.currentToken = ["#"];
		this.offset++;
	}

	private handleComma(): void {
		this.finalizeCurrentToken();

		// Check if inside an array
		const inArray = this.stack.some(
			(t) => t.value === "ARRAY" || t.value === "ARRAYROW",
		);
		const inFunction = this.stack.some(
			(t) => t.type === TokenType.Function,
		);

		if (inArray) {
			// Treat comma as column separator within array
			this.addToken(TokenType.OperatorInfix, TokenSubType.Union, ",");
		} else if (inFunction) {
			this.addToken(TokenType.Argument, "", ",");
		} else {
			this.addToken(TokenType.OperatorInfix, TokenSubType.Union, ",");
		}

		this.offset++;
	}

	private handleColon(): void {
		this.finalizeCurrentToken();
		this.addToken(TokenType.OperatorInfix, TokenSubType.Range, ":");
		this.offset++;
	}

	private handleArraySeparator(): void {
		this.finalizeCurrentToken();
		// Only add ARRAYROW for semicolons that aren't inside a nested array
		const lastArrayDepth = this.stack.filter(
			(t) => t.value === "ARRAY",
		).length;
		if (lastArrayDepth === 1) {
			// We're at the top level array
			this.addToken(TokenType.Function, TokenSubType.Stop, "ARRAYROW");
			this.addToken(TokenType.Argument, "", ";");
			this.stack.push(
				this.addToken(
					TokenType.Function,
					TokenSubType.Start,
					"ARRAYROW",
				),
			);
		} else {
			// We're in a nested array, just add the separator
			this.addToken(TokenType.Argument, "", ";");
		}
		this.offset++;
	}

	private resetState(): void {
		this.tokens = [];
		this.stack = [];
		this.offset = 0;
		this.currentToken = [];
		this.inString = false;
		this.inPath = false;
		this.inRange = false;
		this.inError = false;
	}

	private processNextCharacter(): void {
		const char = this.chars[this.offset];

		if (this.handleComplexStates(char)) return;
		if (this.handleOperators(char)) return;
		if (this.handleStructuralCharacters(char)) return;
		if (this.handleWhitespace(char)) return;

		this.currentToken.push(char);
		this.offset++;
	}

	private handleComplexStates(char: string): boolean {
		if (this.inString) return this.handleString(char);
		if (this.inPath) return this.handlePath(char);
		if (this.inRange) return this.handleRange(char);
		if (this.inError) return this.handleError(char);
		return false;
	}

	private handleString(char: string): boolean {
		if (char === '"') {
			if (this.chars[this.offset + 1] === '"') {
				this.currentToken.push('"');
				this.offset++;
			} else {
				this.inString = false;
				this.addToken(TokenType.Operand, TokenSubType.Text);
			}
		} else {
			this.currentToken.push(char);
		}
		this.offset++;
		return true;
	}

	private handlePath(char: string): boolean {
		if (char === "'") {
			if (this.chars[this.offset + 1] === "'") {
				this.currentToken.push("'");
				this.offset++;
			} else {
				this.inPath = false;
			}
		} else {
			this.currentToken.push(char);
		}
		this.offset++;
		return true;
	}

	private handleRange(char: string): boolean {
		if (char === "]") this.inRange = false;
		this.currentToken.push(char);
		this.offset++;
		return true;
	}

	private handleError(char: string): boolean {
		// Continue consuming characters until we hit a delimiter
		while (this.offset < this.chars.length) {
			const currentChar = this.chars[this.offset];
			// Stop if we hit an operator or whitespace
			if (/[\s+\-*/^&=<>,()]/.test(currentChar)) {
				break;
			}
			this.currentToken.push(currentChar);
			this.offset++;
		}

		const error = this.currentToken.join("");
		this.inError = false;
		this.addToken(TokenType.Operand, TokenSubType.Error);

		return true;
	}

	private isExcelError(value: string): boolean {
		return value.startsWith("#");
	}

	private handleOperators(char: string): boolean {
		const nextChar = this.chars[this.offset + 1] || "";
		const doubleChar = char + nextChar;

		if (this.isComparisonOperator(doubleChar)) {
			this.addTokenFromBuffer();
			this.addToken(
				TokenType.OperatorInfix,
				TokenSubType.Logical,
				doubleChar,
			);
			this.offset += 2;
			return true;
		}

		if (this.isInfixOperator(char)) {
			this.addTokenFromBuffer();
			this.addToken(TokenType.OperatorInfix, "", char);
			this.offset++;
			return true;
		}

		if (char === "%") {
			this.addTokenFromBuffer();
			this.addToken(TokenType.OperatorPostfix, "", char);
			this.offset++;
			return true;
		}

		return false;
	}

	private isComparisonOperator(op: string): boolean {
		return ["<=", ">=", "<>"].includes(op);
	}

	private isInfixOperator(char: string): boolean {
		return "+-*/^&=><".includes(char);
	}

	private handleStructuralCharacters(char: string): boolean {
		switch (char) {
			case "(":
				this.handleOpenParen();
				return true;
			case ")":
				this.handleCloseParen();
				return true;
			case "{":
				this.handleOpenBrace();
				return true;
			case "}":
				this.handleCloseBrace();
				return true;
			case "[":
				this.inRange = true;
				this.currentToken.push(char);
				this.offset++;
				return true;
			case "#":
				this.handleHash();
				return true;
			case ",":
				this.handleComma();
				return true;
			case ":":
				this.handleColon();
				return true;
			case ";":
				this.handleArraySeparator();
				return true;
		}
		return false;
	}

	private handleOpenParen(): void {
		this.addTokenFromBuffer();

		// Check if previous token is a potential function name
		const prevToken = this.tokens[this.tokens.length - 1];
		if (prevToken && prevToken.type === TokenType.Operand) {
			// Convert operand to function
			prevToken.type = TokenType.Function;
			prevToken.subtype = TokenSubType.Start;
			this.stack.push(prevToken);
		} else {
			// Handle as subexpression
			this.stack.push(
				this.addToken(TokenType.Subexpression, TokenSubType.Start, ""),
			);
		}
		this.offset++;
	}

	private handleCloseParen(): void {
		this.addTokenFromBuffer();

		let tokenAdded = false;
		if (this.stack.length > 0) {
			const stackToken = this.stack.pop();
			if (stackToken?.type === TokenType.Function) {
				this.addToken(TokenType.Function, TokenSubType.Stop, "");
				tokenAdded = true;
			} else {
				this.addToken(TokenType.Subexpression, TokenSubType.Stop, "");
				tokenAdded = true;
			}
		}
		// Add a closing parenthesis even if there's no matching opening
		if (!tokenAdded) {
			this.addToken(TokenType.Subexpression, TokenSubType.Stop, "");
		}
		this.offset++;
	}

	public render(tokens: FormulaToken[]): string {
		let output = "";
		let arrayDepth = 0;

		for (const token of tokens) {
			switch (token.type) {
				case TokenType.Function:
					if (token.value === "ARRAY" || token.value === "ARRAYROW") {
						if (token.subtype === TokenSubType.Start) {
							output += "{";
							arrayDepth++;
						} else {
							output += "}";
							arrayDepth--;
						}
					} else {
						// Normal function handling
						if (token.subtype === TokenSubType.Start) {
							output += `${token.value}(`;
						} else {
							output += ")";
						}
					}
					break;

				case TokenType.Subexpression:
					output += token.subtype === TokenSubType.Start ? "(" : ")";
					break;

				case TokenType.OperatorInfix:
					if (token.value === ",") {
						output += arrayDepth > 0 ? "," : ",";
					} else {
						output += ` ${token.value} `;
					}
					break;

				case TokenType.OperatorPostfix:
					output += token.value;
					break;

				case TokenType.Argument:
					output += arrayDepth > 1 ? ";" : ",";
					break;

				case TokenType.Operand:
					if (token.subtype === TokenSubType.Text) {
						output += `"${token.value}"`;
					} else {
						output += token.value;
					}
					break;

				default:
					output += token.value;
			}
		}

		return output
			.replace(/\s+/g, " ")
			.replace(/([+\-*/^=<>])\s+/g, "$1")
			.replace(/\s+([+\-*/^=<>])/g, "$1")
			.replace(/" "/g, '""')
			.trim();
	}
	private handleWhitespace(char: string): boolean {
		if (char !== " ") return false;

		this.addTokenFromBuffer();
		this.addToken(TokenType.Whitespace, "", " ");
		this.offset++;

		// Skip consecutive whitespaces
		while (this.chars[this.offset] === " ") {
			this.offset++;
		}
		return true;
	}

	private addTokenFromBuffer(): void {
		if (this.currentToken.length === 0) return;

		const value = this.currentToken.join("");
		this.currentToken = [];

		if (this.isNumber(value)) {
			this.addToken(TokenType.Operand, TokenSubType.Number, value);
		} else if (this.isLogical(value)) {
			this.addToken(TokenType.Operand, TokenSubType.Logical, value);
		} else {
			this.addToken(TokenType.Operand, TokenSubType.Range, value);
		}
	}

	private isNumber(value: string): boolean {
		return (
			!Number.isNaN(Number.parseFloat(value)) &&
			Number.isFinite(Number(value))
		);
	}

	private isLogical(value: string): boolean {
		return value === "TRUE" || value === "FALSE";
	}

	private addToken(
		type: TokenType,
		subtype: TokenSubType | "",
		value?: string,
	): FormulaToken {
		const tokenValue = value || this.currentToken.join("");
		const token: FormulaToken = {
			value: tokenValue,
			type,
			subtype,
		};

		if (type === TokenType.Function && tokenValue === "(") {
			const prevToken = this.tokens[this.tokens.length - 1];
			if (prevToken && prevToken.type === TokenType.Operand) {
				prevToken.type = TokenType.Function;
				prevToken.subtype = TokenSubType.Start;
				this.stack.push(prevToken);
			}
		} else {
			this.tokens.push(token);
		}

		this.currentToken = [];
		return token;
	}

	private postProcessTokens(): FormulaToken[] {
		return this.tokens
			.filter((token) => token.type !== TokenType.Whitespace)
			.map((token) => {
				// Clean up function markers
				if (token.type === TokenType.Function && token.value === "(") {
					return null;
				}
				return this.adjustTokenTypes(token);
			})
			.filter((token): token is FormulaToken => token !== null);
	}

	private adjustTokenTypes(token: FormulaToken): FormulaToken {
		if (token.type === TokenType.OperatorInfix) {
			if (token.value === "&") {
				token.subtype = TokenSubType.Concatenation;
			} else if (["=", "<", ">"].includes(token.value)) {
				token.subtype = TokenSubType.Logical;
			} else {
				token.subtype = TokenSubType.Math;
			}
		}

		if (
			token.type === TokenType.Operand &&
			token.subtype === TokenSubType.Range
		) {
			if (token.value.startsWith("#")) {
				token.subtype = TokenSubType.Error;
			}
		}

		return token;
	}

	public prettyPrint(tokens: FormulaToken[]): string {
		let output = "";
		let indent = 0;

		for (const token of tokens) {
			if (token.subtype === TokenSubType.Stop)
				indent = Math.max(0, indent - 1);

			output += "  ".repeat(indent);
			output += `${token.value} <${token.type}> <${token.subtype}>\n`;

			if (token.subtype === TokenSubType.Start) indent++;
		}

		return output;
	}
}

/* Usage example:
const parser = new ExcelFormulaParser();
const tokens = parser.parse("=SUM(A3+B9*2)/2");
console.log(parser.prettyPrint(tokens));
*/
