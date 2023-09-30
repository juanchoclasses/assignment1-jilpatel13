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
    this._errorMessage = "";
    this._errorOccured = false;

    // set the this._result to the length of the formula
    this._currentFormula = [...formula];

    // if the formula is empty return 0
    if (formula.length === 0) {
      this._result = 0;
      this._errorMessage = ErrorMessages.emptyFormula;
      return;
    }

    let result = this.addition();
    this._result = result;

    if (!this._errorOccured && this._currentFormula.length > 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }

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

  // function to add a number to the result
  private addition(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = this.multiplication();
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "-" || this._currentFormula[0] === "+")
    ) {
      let token = this._currentFormula.shift();
      let multiplication = this.multiplication();
      if (token === "+") {
        result += multiplication;
      } else {
        result -= multiplication;
      }
    }
    this._lastResult = result;
    return result;
  }

  // function to multiply a number to the result
  private multiplication(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = this.parentheses();
    while (
      this._currentFormula.length > 0 &&
      (this._currentFormula[0] === "*" || this._currentFormula[0] === "/")
    ) {
      let token = this._currentFormula.shift();
      let parentheses = this.parentheses();
      if (token === "*") {
        result *= parentheses;
      } else if (token === "/") {
        if (parentheses === 0) {
          this._errorOccured = true;
          this._errorMessage = ErrorMessages.divideByZero;
          this._lastResult = Infinity;
          return Infinity;
        }
        result /= parentheses;
      }
    }
    this._lastResult = result;
    return result;
  }

  // function to evaluate parentheses
  private parentheses(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = 0;
    if (this._currentFormula.length === 0) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.partial;
      return result;
    }

    let token = this._currentFormula.shift();

    if (this.isNumber(token)) {
      result = Number(token);
      this._lastResult = result;
    } else if (token === "(") {
      result = this.addition();
      if (
        this._currentFormula.length === 0 ||
        this._currentFormula.shift() !== ")"
      ) {
        this._errorOccured = true;
        this._errorMessage = ErrorMessages.missingParentheses;
        this._lastResult = result;
      }
    } else if (this.isCellReference(token)) {
      [result, this._errorMessage] = this.getCellValue(token);

      if (this._errorMessage !== "") {
        this._errorOccured = true;
        this._lastResult = result;
      }
    } else {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
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
