/** Value and error primitives shared by the evaluator and the function library. */

export class FormulaError {
  code: string
  constructor(code: string) { this.code = code }
  toString() { return this.code }
}

export type Scalar = number | string | boolean | FormulaError | null

/** Excel's text rendering of a value (used by & concatenation and the text functions). */
export const toStr = (v: Scalar): string =>
  v === null ? ''
    : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE')
      : typeof v === 'number' ? String(Number(v.toPrecision(15)))
        : String(v)
