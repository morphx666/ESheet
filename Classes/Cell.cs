using ESheet.Classes;

internal class Cell(Sheet sheet, int col, int row) {
    private string value = "";
    private double valueEvaluated = 0;
    private readonly Sheet sheet = sheet;

    public enum Types {
        Empty,
        Number,
        Label,
        Formula
    }

    public enum Alignments {
        Left,
        Center,
        Right
    }

    public Types Type { get; private set; }

    public int Column { get; init; } = col;
    public int Row { get; init; } = row;

    public List<Cell> AffectedCells = [];

    public ConsoleColor ForeColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor ForeSelColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor BackSelColor { get; set; } = ConsoleColor.Cyan;

    public Alignments Alignment { get; set; }

    public static Evaluator Eval = new();

    public string Value {
        get => value;
        set {
            this.value = value;
            if(value != "") {
                switch(value[0]) {
                    case '\'':
                        Type = Types.Label;
                        this.value = this.value[1..];
                        Alignment = Alignments.Left;
                        break;
                    case '=':
                        Type = Types.Formula;
                        this.value = this.value[1..];
                        Alignment = Alignments.Right;

                        Eval.Formula = ExpandRanges(this.value);
                        valueEvaluated = Evaluate();
                        break;
                    default:
                        UpdateType();
                        break;
                }
            }
        }
    }

    public string ValueFormat {
        get {
            string f = "";
            switch(Type) {
                case Types.Label:
                    f = "'";
                    break;
                case Types.Formula:
                    f = "=";
                    break;
                case Types.Number:
                    break;
            }

            return f + Value;
        }
    }

    public double ValueEvaluated {
        get => valueEvaluated;
    }

    public Cell(Sheet sheet, int col, int row, string value, Alignments alignment) : this(sheet, col, row) {
        this.Value = value.Trim();
        this.Alignment = alignment;
    }

    public Cell(Sheet sheet, int col, int row, string value) : this(sheet, col, row) {
        this.Value = value.Trim();
    }

    private void UpdateType() {
        if(value == "") {
            Type = Types.Empty;
        } else {
            if(double.TryParse(value, out valueEvaluated)) {
                Type = Types.Number;
                Alignment = Alignments.Right;
            } else {
                Eval.Formula = ExpandRanges(value);
                try {
                    valueEvaluated = Evaluate();
                    Type = Types.Formula;
                    Alignment = Alignments.Right;
                } catch {
                    Type = Types.Label;
                    Alignment = Alignments.Left;
                }
            }
        }
    }

    public void Refresh() {
        this.Value = this.value;
    }

    private string ExpandRanges(string formula) {
        string rangeSeparator = "..";

        while(formula.Contains(rangeSeparator)) {
            int index = formula.IndexOf(rangeSeparator);

            string startCell = "";
            for(int i = index - 1; i >= 0; i--) {
                if(char.IsAsciiLetterOrDigit(formula[i])) {
                    startCell = formula[i] + startCell;
                } else {
                    break;
                }
            }

            string endCell = "";
            for(int i = index + rangeSeparator.Length; i < formula.Length; i++) {
                if(char.IsAsciiLetterOrDigit(formula[i])) {
                    endCell += formula[i];
                } else {
                    break;
                }
            }

            (int Column, int Row) start = sheet.GetCellColRow(startCell);
            (int Column, int Row) end = sheet.GetCellColRow(endCell);
            string range = "";

            for(int row = Math.Min(start.Row, end.Row); row <= Math.Max(start.Row, end.Row); row++) {
                for(int col = Math.Min(start.Column, end.Column); col <= Math.Max(start.Column, end.Column); col++) {
                    range += sheet.GetCellName(col, row) + ",";
                }
            }
            range = range.TrimEnd(',');
            formula = formula.Replace($"{startCell}..{endCell}", range);
        }

        return formula;
    }    

    private double Evaluate() {
        double res = 0;
        AffectedCells.Clear();

        Eval.CustomParameters.Clear();
        while(true) {
            try {
                res = (double)Eval.Evaluate();
                break;
            } catch(ArgumentException ex) when(ex.ParamName is not null) {
                string name = ex.ParamName;
                Cell? cell = sheet.GetCell(name);
                if(cell == null) {
                    //TODO: Do something here. Return 0?
                    break;
                }

                if(cell.Type == Types.Label) {
                    //TODO: Do something here, b/c we cannot have Label type cells as part of a formula
                    //      Unless we add string manipulation functions??? hmmm...
                    break;
                }
                double value = cell.ValueEvaluated;
                Eval.CustomParameters.Add(name, value);
                AffectedCells.Add(cell);
            }
        }

        return res;
    }
}