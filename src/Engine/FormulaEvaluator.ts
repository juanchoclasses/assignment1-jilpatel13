import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";

export class FormulaEvaluator {
  // Define a function called update that takes a string parameter and returns a number
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  evaluate(formula: FormulaType) {
    // set the current formula to the new formula
    this._currentFormula = [...formula];
    this._lastResult = 0;
    // clear the error message
    this._errorMessage = "";
    // set the errorOccured flag
    this._errorOccured = false;

    switch (formula.length) {
      case 0:
        this._errorMessage = ErrorMessages.emptyFormula;
        break;
      default:
        this._errorMessage = "";
        break;
    }

    // if the formula is empty return set the result to 0 and return
    if (this._errorMessage === ErrorMessages.emptyFormula) {
      this._result = 0;
      return;
    }

    let result = this.calculate();
    this._result = result;

    // if there was an error set the result to invalid formula
    if (this._currentFormula.length > 0 && !this._errorOccured) {
      this._errorMessage = ErrorMessages.invalidFormula;
      this._errorOccured = true;
    }

    // if there was an error set the result to the last result
    if (this._errorOccured) {
      this._result = this._lastResult;
    }
  }

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }

  // calculate the result of the formula
  private calculate(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    // get the first token
    let result = this.token();
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "+" || this._currentFormula[0] === "-")
    ) {
      let operator = this._currentFormula.shift();
      let token = this.token();
      if (operator === "+") {
        result += token;
      } else if (operator === "-") {
        result -= token;
      }
    }
    // set the lastResult to the result
    this._lastResult = result;
    return result;
  }

  private operator(): boolean {
    switch (this._currentFormula[0]) {
      case "/":
      case "*":
      case "+/-":
        return true;
      default:
        return false;
    }
  }

  private token(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = this.parse();
    while (this._currentFormula.length > 0 && this.operator()) {
      let operator = this._currentFormula.shift();

      if (this._errorOccured && operator === "+/-") {
        this._errorMessage = "";
        this._errorOccured = false;
      }

      // if the operator is +/-
      if (operator === "+/-") {
        if (result === 0) {
          result = 0;
        } else {
          result = result * -1;
        }
        continue;
      }

      let number = this.parse();
      if (operator === "*") {
        result *= number;
      } else if (operator === "/") {
        // check for divide by zero
        if (number === 0) {
          this._errorMessage = ErrorMessages.divideByZero;
          this._errorOccured = true;
          this._lastResult = Infinity;
          return Infinity;
        } else {
          result /= number;
        }
      }
    }
    // set the lastResult to the result
    this._lastResult = result;
    return result;
  }

  private parse(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = 0;

    if (this._currentFormula.length === 0) {
      this._errorMessage = ErrorMessages.partial;
      this._errorOccured = true;
      return result;
    }

    // get the current token
    let token = this._currentFormula.shift();

    // if the token is a number set the result to the number
    if (this.isNumber(token)) {
      result = Number(token);
      this._lastResult = result;
    } else if (token === "(") {
      result = this.calculate();
      if (
        this._currentFormula.length === 0 ||
        this._currentFormula.shift() !== ")"
      ) {
        this._errorMessage = ErrorMessages.missingParentheses;
        this._errorOccured = true;
        this._lastResult = result;
      }
    } else if (this.isCellReference(token)) {
      [result, this._errorMessage] = this.getCellValue(token);

      if (this._errorMessage !== "") {
        this._errorOccured = true;
        this._lastResult = result;
      }
    } else {
      this._errorMessage = ErrorMessages.invalidFormula;
      this._errorOccured = true;
    }
    return result;
  }
  /**
   *
   * @param token
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  /**
   *
   * @param token
   * @returns true if the token is a cell reference
   *
   */
  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  /**
   *
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   *
   */
  getCellValue(token: TokenType): [number, string] {
    let cell = this._sheetMemory.getCellByLabel(token);
    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error !== "" && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    let value = cell.getValue();
    return [value, ""];
  }
}

export default FormulaEvaluator;
