using System.Drawing;
using System.Net.Http.Headers;

internal partial class Sheet {
    public void Render() {
        int sc = workingMode == Modes.Formula ? SelFormulaColumn : SelColumn;
        int sr = workingMode == Modes.Formula ? SelFormulaRow : SelRow;
        if(SheetColumnToConsoleColumn(sc - StartColumn) >= Console.WindowWidth - OffsetLeft) StartColumn++;
        if(sc < StartColumn) StartColumn--;
        if((sr - StartRow) >= Console.WindowHeight - OffsetTop - 1) StartRow++;
        if(sr < StartRow) StartRow--;

        Console.CursorVisible = false;

        SetColors(BackHeaderColor, BackCellColor);
        Console.SetCursorPosition(0, 0);
        Console.Write($"{GetColumnName(sc)}{sr + 1}: ");

        SetColors(ConsoleColor.White, ConsoleColor.Black);
        WriteLine(GetCell(sc, sr)?.ValueFormat ?? "");

        RenderHeaders();
        RenderSheet();

        Console.CursorVisible = true;
    }

    private static void RenderHelp(int c, int r, string title, (string Key, string Action)[] values) {
        Console.SetCursorPosition(c, r);

        SetColors(ConsoleColor.Yellow, ConsoleColor.Black);
        Console.Write($"{title}: ");

        int count = values.Length;
        foreach((string Key, string Action) in values) {
            SetColors(ConsoleColor.DarkGray, ConsoleColor.Black);
            Console.Write("[");

            SetColors(ConsoleColor.DarkGreen, ConsoleColor.Black);
            Console.Write($"{Key}");

            SetColors(ConsoleColor.DarkGray, ConsoleColor.Black);
            Console.Write("] ");

            SetColors(ConsoleColor.DarkYellow, ConsoleColor.Black);
            Console.Write($"{Action}");

            if(--count > 0) {
                SetColors(ConsoleColor.DarkGray, ConsoleColor.Black);
                Console.Write(" | ");
            }
        }

        SetColors(ConsoleColor.Black, ConsoleColor.Black);
        Console.Write(" ".PadLeft(Console.WindowWidth - Console.CursorLeft));
    }

    private void RenderHeaders() {
        SetColors(ConsoleColor.DarkGray, BackCellColor);

        if(lastWorkingMode != workingMode) {
            RenderHelp(OffsetLeft, OffsetTop - 1, WorkingModeToString(), helpMessages[WorkingModeToKey()]);
            lastWorkingMode = workingMode;
        }

        SetColors(ForeHeaderColor, BackHeaderColor);
        Console.SetCursorPosition(OffsetLeft, OffsetTop);
        Console.Write(AlignText(" ", RowWidth, Cell.Alignments.Left));

ReStart:
        for(int r = 1; r < Console.WindowHeight - OffsetTop; r++) {
            if(r - 1 == SelRow - StartRow) {
                SetColors(ForeHeaderColor, BackHeaderSelColor);
            } else {
                SetColors(ForeHeaderColor, BackHeaderColor);
            }
            Console.SetCursorPosition(OffsetLeft, OffsetTop + r);

            string rowLabel = (r + StartRow).ToString();
            if(rowLabel.Length >= RowWidth) {
                RowWidth = rowLabel.Length + 1;
                goto ReStart; // Another goto? Am I going mad??? Well, it's better than recursion...
            }
            Console.Write(AlignText(rowLabel, RowWidth, Cell.Alignments.Right));
        }

        int c = 0;
        while(true) {
            if(c == SelColumn - StartColumn) {
                SetColors(ForeHeaderColor, BackHeaderSelColor);
            } else {
                SetColors(ForeHeaderColor, BackHeaderColor);
            }
            int cc = SheetColumnToConsoleColumn(c);
            Console.SetCursorPosition(OffsetLeft + cc, OffsetTop);
            (string Text, bool Overflow) result = Trim(AlignText(GetColumnName(c + StartColumn), GetColumnWidth(c + StartColumn), Cell.Alignments.Center), cc);
            Console.Write(result.Text);

            if(result.Overflow) break;
            c++;
        }
    }

    private void RenderSheet() {
        int sc = workingMode == Modes.Formula ? SelFormulaColumn : SelColumn;
        int sr = workingMode == Modes.Formula ? SelFormulaRow : SelRow;
        int col = 0;
        int row = 0;
        int cc;
        (string Text, bool Overflow) result;
        List<Cell> dependentCells = [];

        while(true) {
            Cell? cell = GetCell(col + StartColumn, row + StartRow);

            cc = SheetColumnToConsoleColumn(col);
            if((col == SelColumn - StartColumn) && (row == SelRow - StartRow)) {
                SetColors(ForeSelCellColor, BackSelCellColor);
            } else if((workingMode == Modes.Formula) && (col == SelFormulaColumn - StartColumn) && (row == SelFormulaRow - StartRow)) {
                SetColors(ConsoleColor.White, ConsoleColor.Red);
            } else {
                if(cell?.HasError ?? false) {
                    SetColors(ConsoleColor.Red, BackCellColor);
                } else {
                    SetColors(ForeCellColor, BackCellColor);
                }
            }

            string emptyCell = Column.GetEmptyCell(this, col + StartColumn);
            if(cell == null) {
                result = Trim(emptyCell, cc);
            } else {
                string value;

                if(cell.HasError) {
                    value = "#ERROR";
                } else {
                    switch(cell.Type) {
                        case Cell.Types.Number:
                            value = cell.ValueEvaluated.Val.ToString($"N{RenderPrecision}");
                            break;
                        case Cell.Types.Formula:
                            if(cell.ValueEvaluated.Str is not null) {
                                value = cell.ValueEvaluated.Str.Replace("\"", "");
                            } else {
                                value = cell.ValueEvaluated.Val.ToString($"N{RenderPrecision}");
                            }
                            break;
                        default:
                            value = cell.Value;
                            break;
                    }
                }
                result = Trim(AlignText(value, Math.Max(value.Length, emptyCell.Length), cell.Alignment), cc);

                if(workingMode != Modes.Formula && (col == sc - StartColumn) && (row == sr - StartRow)) {
                    dependentCells.AddRange(cell.DependentCells);

                    foreach(Cell ac in cell.DependentCells) {
                        // TODO: This is the same code as above... extract it as a method or anonymous function
                        if(ac.HasError) {
                            value = "#ERROR";
                        } else {
                            switch(ac.Type) {
                                case Cell.Types.Number:
                                    value = ac.ValueEvaluated.Val.ToString($"N{RenderPrecision}");
                                    break;
                                case Cell.Types.Formula:
                                    if(ac.ValueEvaluated.Str is not null) {
                                        value = ac.ValueEvaluated.Str.Replace("\"", "");
                                    } else {
                                        value = ac.ValueEvaluated.Val.ToString($"N{RenderPrecision}");
                                    }
                                    break;
                                default:
                                    value = ac.Value;
                                    break;
                            }
                        }
                        string acResult = Trim(AlignText(value, Math.Max(value.Length, Column.GetColumnWidth(this, ac.Column)), cell.Alignment), cc).Text;

                        ConsoleColor fc = Console.ForegroundColor;
                        ConsoleColor bc = Console.BackgroundColor;

                        SetColors(fc, ConsoleColor.DarkGray);
                        Console.SetCursorPosition(SheetColumnToConsoleColumn(ac.Column), OffsetTop + ac.Row + 1);
                        Console.Write(acResult);

                        SetColors(fc, bc);
                    }
                }
            }

            if(cell is not null && dependentCells.Contains(cell)) SetColors(ForeCellColor, ConsoleColor.DarkGray);

            Console.SetCursorPosition(OffsetLeft + cc, OffsetTop + row + 1);
            Console.Write(result.Text);

            if((col == SelColumn - StartColumn) && (row == SelRow - StartRow) && (cell?.HasError ?? false)) {
                SetColors(ConsoleColor.Red, BackCellColor);
                Console.SetCursorPosition(Console.WindowWidth - cell.ErrorMessage.Length - 1, 0);
                Console.Write(cell.ErrorMessage);
            }

            if(result.Overflow) {
                col = 0;
                row++;
                if(row == Console.WindowHeight - OffsetTop - 1) break;
            } else {
                col++;
            }
        }
    }

    private void WriteLine(string text) {
        string emptyRow = new(' ', Console.WindowWidth - OffsetLeft - text.Length);
        Console.Write(text + emptyRow);
    }

    private static void SetColors(ConsoleColor fore, ConsoleColor back) {
        Console.ForegroundColor = fore;
        Console.BackgroundColor = back;
    }
}