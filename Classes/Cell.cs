using ESheet.Classes;
using System.Diagnostics;

internal class Cell {
    private string value = "";
    private double valueEvaluated = 0;
    private Sheet sheet;

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

    public int Column { get; init; }
    public int Row { get; init; }

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

                        Eval.Formula = this.value;
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

    public Cell(Sheet sheet, int col, int row, string value = "", Alignments alignment = Alignments.Left) {
        this.sheet = sheet;
        this.Column = col;
        this.Row = row;
        this.Value = value.Trim();
        this.Alignment = alignment;
    }

    private void UpdateType() {
        if(value == "") {
            Type = Types.Empty;
        } else {
            if(double.TryParse(value, out valueEvaluated)) {
                Type = Types.Number;
                Alignment = Alignments.Right;
            } else {
                Eval.Formula = value;
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

    private double Evaluate() {
        double res = 0;

        Eval.CustomParameters.Clear();
        while(true) {
            try {
                res = (double)Eval.Evaluate();
                break;
            } catch(ArgumentException ex) {
                string name = ex.ParamName;
                double value = sheet.GetCell(name).ValueEvaluated;
                Eval.CustomParameters.Add(name, value);
            }
        }

        return res;
    }
}