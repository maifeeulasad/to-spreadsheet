import { IRows } from ".";

const indexToVbIndex = (index: number) => index + 1;

const indexToVbRelationIndex = (index: number) => indexToVbIndex(index) + 2;

// 1->a, 2->b, 26->z, 27->aa
const indexToRowIndex = (index: number): string => {
    const base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    let result = "";
    do {
        const remainder = index % 26;
        result = base[(remainder || 26) - 1] + result;
        index = Math.floor(index / 26);
    } while (index > 0);
    return result;
}

const rowColumnToVbPosition = (row: number, col: number): string => indexToRowIndex(indexToVbIndex(row)) + indexToVbIndex(col);

const calculateExtant = (rows: IRows[]): string => rowColumnToVbPosition(
    Math.max(...rows.map(row => row.cells.length)) - 1,
    rows.length - 1
)

class SkipCell {
    private skipCell: number;
    public getSkipCell = () => this.skipCell;
    constructor(skipCell: number) {
        this.skipCell = skipCell;
    }
}

const skipCell = (skipCell: number) => new SkipCell(skipCell);

class Equation {
    private equation: string;
    public getEquation = () => this.equation;
    constructor(equation: string) {
        this.equation = equation;
    }
}

const writeEquation = (equation: string) => new Equation(equation);

export { indexToVbIndex, indexToVbRelationIndex, indexToRowIndex, rowColumnToVbPosition, calculateExtant, SkipCell, skipCell, Equation, writeEquation }