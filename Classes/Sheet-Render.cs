internal partial class Sheet {
    public void Render() {
        if(SheetColumnToConsoleColumn(SelColumn - StartColumn) >= Console.WindowWidth - OffsetLeft) StartColumn++;
        if(SelColumn < StartColumn) StartColumn--;
        if((SelRow - StartRow) >= Console.WindowHeight - OffsetTop - 1) StartRow++;
        if(SelRow < StartRow) StartRow--;

        Console.CursorVisible = false;

        RenderHeaders();
        RenderSheet();

        SetColors(BackHeaderColor, BackCellColor);
        Console.SetCursorPosition(0, 0);
        Console.Write($"{GetColumnName(SelColumn)}{SelRow + 1}: ");

        SetColors(ConsoleColor.White, ConsoleColor.Black);
        WriteLine(GetCell(SelColumn, SelRow)?.ValueFormat ?? "");

        Console.CursorVisible = true;
    }

    private static void RenderHelp(int c, int r, string title, (string Key, string Action)[] values) {
        Console.SetCursorPosition(c, r);
        SetColors(ConsoleColor.Black, ConsoleColor.Black);
        Console.Write(" ".PadLeft(Console.WindowWidth));
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
    }

    private void RenderHeaders() {
        SetColors(ConsoleColor.DarkGray, BackCellColor);

        RenderHelp(OffsetLeft, OffsetTop - 1, WorkingModeToString(), helpMessages[WorkingModeToKey()]);

        SetColors(ForeHeaderColor, BackHeaderColor);
        Console.SetCursorPosition(OffsetLeft, OffsetTop);
        Console.Write(AlignText(" ", RowWidth, Cell.Alignments.Left));

        for(int r = 1; r < Console.WindowHeight - OffsetTop; r++) {
            if(r - 1 == SelRow - StartRow) {
                SetColors(ForeHeaderColor, BackHeaderSelColor);
            } else {
                SetColors(ForeHeaderColor, BackHeaderColor);
            }
            Console.SetCursorPosition(OffsetLeft, OffsetTop + r);
            Console.Write(AlignText((r + StartRow).ToString(), RowWidth, Cell.Alignments.Left));
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
            (string Text, bool Overflow) result = Trim(AlignText(GetColumnName(c + StartColumn), CellWidth, Cell.Alignments.Center), cc);
            Console.Write(result.Text);

            if(result.Overflow) break;
            c++;
        }
    }

    private void RenderSheet() {
        int col = 0;
        int row = 0;
        int cc;
        (string Text, bool Overflow) result;
        List<Cell> dependentCells = [];

        while(true) {
            cc = SheetColumnToConsoleColumn(col);
            if((col == SelColumn - StartColumn) && (row == SelRow - StartRow)) {
                SetColors(ForeSelCellColor, BackSelCellColor);
            } else {
                if((workingMode == Modes.Formula) && col == SelFormulaColumn && row == SelFormulaRow) {
                    SetColors(ConsoleColor.White, ConsoleColor.Red);
                } else {
                    SetColors(ForeCellColor, BackCellColor);
                }
            }

            Cell? cell = GetCell(col + StartColumn, row + StartRow);
            if(cell == null) {
                //if((c == SelColumn - StartColumn) && (r == SelRow - StartRow)) {
                //    result = Trim(AlignText($"{c + StartColumn}:{r + StartRow}", CellWidth, Cell.Alignments.Center), cc);
                //} else {
                //    result = Trim(emptyCell, cc);
                //}
                result = Trim(emptyCell, cc);
            } else {
                string value;
                if(cell.Type == Cell.Types.Number || cell.Type == Cell.Types.Formula) {
                    value = cell.ValueEvaluated.ToString("N2");
                } else {
                    value = cell.Value;
                }
                result = Trim(AlignText(value, Math.Max(value.Length, emptyCell.Length), cell.Alignment), cc);

                if(workingMode != Modes.Formula && (col == SelColumn - StartColumn) && (row == SelRow - StartRow)) {
                    dependentCells.AddRange(cell.DependentCells);

                    foreach(Cell ac in cell.DependentCells) {
                        // TODO: This is the same code as above... extract it as a method or anonymous function
                        if(ac.Type == Cell.Types.Number || ac.Type == Cell.Types.Formula) {
                            value = ac.ValueEvaluated.ToString("N2");
                        } else {
                            value = ac.Value;
                        }
                        string acResult = Trim(AlignText(value, Math.Max(value.Length, emptyCell.Length), cell.Alignment), cc).Text;

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