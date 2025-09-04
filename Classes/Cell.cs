using System.Diagnostics;
using System.Linq;

internal class Cell(Sheet sheet, int col, int row) {
    private string value = "";
    private (double Val, string? Str) valueEvaluated = (0, null);
    private readonly Sheet sheet = sheet;
    private bool hasError = false;
    private string errorMessage = "";

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

    public int Column { get; private set; } = col;
    public int Row { get; private set; } = row;

    public List<Cell> DependentCells = [];

    public ConsoleColor ForeColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor ForeSelColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor BackSelColor { get; set; } = ConsoleColor.Cyan;

    public Alignments Alignment { get; set; }

    public bool HasError { get => hasError; }
    public string ErrorMessage { get => errorMessage; }

    public static Evaluator Eval = new();

    public string Value {
        get => value;
        set {
            bool hadError = hasError;
            hasError = false;
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

                        Eval.Formula = ExtractStrings(ExpandRanges(this.value));
                        valueEvaluated = Evaluate();
                        break;
                    default:
                        UpdateType();
                        break;
                }
                if(hadError && !hasError) sheet.FullRefresh();
            }
        }
    }

    internal void SetValueFast(string value) {
        this.value = value;
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

    public (double Val, string? Str) ValueEvaluated {
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
            if(double.TryParse(value, out double evalValue)) {
                valueEvaluated = (evalValue, null);
                Type = Types.Number;
                Alignment = Alignments.Right;
            } else {
                Eval.Formula = ExtractStrings(ExpandRanges(this.value));
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

    internal void SetColRow(int col, int row) {
        this.Column = col;
        this.Row = row;
    }

    internal string ExpandRanges(string formula) {
        string rangeSeparator = "..";
        formula = formula.Replace("%", "*1/100");

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

    internal static string ExtractStrings(string formula) {
        Eval.Strings.Clear();

        int p0 = formula.IndexOf('"');
        while(p0 != -1) {
            int p1 = formula.IndexOf('"', p0 + 1);
            if(p1 == -1) {
                // Missing closing quote
                break;
            }

            string str = formula.Substring(p0, p1 - p0 + 1);
            int id = Eval.Strings.Count;
            Eval.Strings.Add(id, str);
            formula = formula.Replace(str, $"STR({id})");

            p0 = formula.IndexOf('"');
        }

        return formula;
    }

    private (double Val, string? Str) Evaluate() {
        (double Val, string? Str) result;
        DependentCells.Clear();

        Eval.CustomParameters.Clear();
        while(true) {
            try {
                result = Eval.Evaluate();
                break;
            } catch(ArgumentException ex) when(ex.ParamName is not null) {
                string name = ex.ParamName;
                if((name.StartsWith('"') && name.EndsWith('"')) ||
                   (name.StartsWith('\'') && name.EndsWith('\''))) {
                    return (0, null);
                }

                Cell? cell = sheet.GetCell(name);
                if(cell == null) {
                    (bool IsValid, int Column, int Row) = sheet.IsCellNameValid(name);
                    if(IsValid) {
                        cell = new(sheet, Column, Row) { Value = "" };
                    } else {
                        SetError($"Unrecognized cell '{name}'");
                        return (0, null);
                    }
                } else if(cell.Type == Types.Label) {
                    SetError($"Invalid cell type '{name}'");
                    return (0, null);
                } else if(cell.HasError) {
                    SetError(ex.Message);
                    return (0, null);
                }

                if(cell.ValueEvaluated.Str is not null) {
                    Eval.Formula = ExtractStrings(cell.ValueEvaluated.Str);
                }
                Eval.CustomParameters.Add(name, cell.ValueEvaluated);
                DependentCells.Add(cell);
            } catch(Exception ex) {
                SetError(ex.Message);
                return (0, null);
            }
        }

        if(Eval.CustomParameters.ContainsKey(sheet.GetCellName(this))) {
            SetError("Circular reference");
            return (0, null);
        }

        // TODO: Should we reset the hasError flag here?
        return result;
    }

    internal void SetError(string message) {
        hasError = true;
        errorMessage = message;
    }

    internal List<(string Name, int Pos, int Column, int Row)> GetReferencedCells() {
        if(Type != Types.Formula) return [];
        List<(string Name, int Pos, int Column, int Row)> cells = [];

        for(int i = 0; i < value.Length; i++) {
            if(char.IsAsciiLetter(value[i])) {
                string cellName = value[i].ToString();
                for(int j = i + 1; j < value.Length; j++) {
                    if(char.IsAsciiLetterOrDigit(value[j])) {
                        cellName += value[j];
                    } else {
                        break;
                    }
                }
                if(char.IsDigit(cellName[^1])) {
                    (bool IsValid, int Column, int Row) = sheet.IsCellNameValid(cellName);
                    if(IsValid) cells.Add((cellName, i, Column, Row));
                }
            }
        }

        return cells;
    }

    public override string ToString() {
        //return $"Cell {sheet.GetCellName(Column, Row)}: Type={Type}, Value='{Value}', ValueEvaluated=({ValueEvaluated.Val}, '{ValueEvaluated.Str}'), HasError={HasError}, ErrorMessage='{ErrorMessage}'";
        return $"Cell ({Row}, {Column}): '{Value}'";
    }
}