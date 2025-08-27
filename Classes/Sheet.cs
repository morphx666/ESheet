using System.Text;

internal class Sheet {
    public int StartColumn { get; set; } = 0;
    public int StartRow { get; set; } = 0;
    public int SelColumn { get; set; } = 0;
    public int SelRow { get; set; } = 0;

    public int OffsetLeft { get; set; } = 0;
    public int OffsetTop { get; set; } = 3;

    public List<Cell> Cells { get; init; } = new();
    public int CellWidth { get; set; } = 15;
    public int RowWidth { get; set; } = 4;

    public ConsoleColor ForeCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackCellColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor ForeHeaderColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor BackHeaderColor { get; set; } = ConsoleColor.Cyan;
    public ConsoleColor ForeSelCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackSelCellColor { get; set; } = ConsoleColor.Blue;

    private readonly int ccCount = 'Z' - 'A' + 1; // 26
    private readonly string emptyCell;

    public Sheet() {
        emptyCell = AlignText(" ", CellWidth, Cell.Alignments.Left);
    }

    public void Render() {
        if(SheetColumnToConsoleColumn(SelColumn - StartColumn) >= Console.WindowWidth - OffsetLeft)
            StartColumn++;
        if(SelColumn < StartColumn)
            StartColumn--;
        if((SelRow - StartRow) >= Console.WindowHeight - OffsetTop - 1)
            StartRow++;
        if(SelRow < StartRow)
            StartRow--;

        Console.CursorVisible = false;

        RenderHeaders();
        RenderSheet();

        Console.SetCursorPosition(0, 0);
        Console.Write($"{GetColumnName(SelColumn)}{SelRow + 1}: ");
        WriteLine(GetCell(SelColumn, SelRow)?.ValueFormat ?? "");

        Console.CursorVisible = true;
    }

    public void Run() {
        string userInput = "";

        while(true) {
            Render();

            Console.SetCursorPosition(0, 1);
            WriteLine(userInput);
            Console.SetCursorPosition(userInput.Length, 1);

            ConsoleKeyInfo ck = Console.ReadKey(true);

            switch(ck.Key) {
                case ConsoleKey.UpArrow:
                    if(SelRow > 0) SelRow--;
                    break;
                case ConsoleKey.DownArrow:
                    SelRow++;
                    break;
                case ConsoleKey.LeftArrow:
                    if(SelColumn > 0) SelColumn--;
                    break;
                case ConsoleKey.RightArrow:
                    SelColumn++;
                    break;
                case ConsoleKey.Escape:
                    return;

                case ConsoleKey.Backspace:
                    if(userInput.Length > 0) userInput = userInput[..^1];
                    break;

                case ConsoleKey.Enter:
                    if(userInput.Length > 0) {
                        Cell? cell = GetCell(SelColumn, SelRow);
                        if(cell == null) {
                            cell = new Cell(this, SelColumn, SelRow);
                            Cells.Add(cell);
                        }

                        cell.Value = userInput;

                        string name = GetCellName(SelColumn, SelRow);
                        CascadeUpdate(name);

                        userInput = "";
                    }
                    break;

                default:
                    if(userInput.Length < Console.WindowWidth - OffsetLeft - RowWidth)
                        userInput += ck.KeyChar;
                    break;
            }
        }
    }

    private void CascadeUpdate(string name) {
        List<string> cellsToUpdate = [];
        foreach(Cell c in Cells) {
            if(c.Type == Cell.Types.Formula && c.Value.Contains(name)) {
                c.Refresh();
                string cName = GetCellName(c);
                cellsToUpdate.Add(cName);
            }
        }

        cellsToUpdate.ForEach(c => CascadeUpdate(c));
    }

    private void WriteLine(string text) {
        string emptyRow = new(' ', Console.WindowWidth - OffsetLeft - text.Length);
        Console.Write(text + emptyRow);
    }

    internal Cell? GetCell(int col, int row) {
        foreach(Cell cell in Cells) {
            if(cell.Column == col && cell.Row == row) {
                return cell;
            }
        }

        return null;
    }

    internal Cell? GetCell(string name) {
        (int col, int row) = GetCellColRow(name);
        return GetCell(col, row);
    }

    private void RenderSheet() {
        int c = 0;
        int r = 0;
        int cc;
        (string Text, bool Overflow) result;

        while(true) {
            cc = SheetColumnToConsoleColumn(c);
            if((c == SelColumn - StartColumn) && (r == SelRow - StartRow)) {
                SetColors(ForeSelCellColor, BackSelCellColor);
            } else {
                SetColors(ForeCellColor, BackCellColor);
            }

            Cell? cell = GetCell(c, r);
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
            }

            Console.SetCursorPosition(OffsetLeft + cc, OffsetTop + r + 1);
            Console.Write(result.Text);

            if(result.Overflow) {
                c = 0;
                r++;
                if(r == Console.WindowHeight - OffsetTop - 1)
                    break;
            } else {
                c++;
            };
        }
    }

    private void RenderHeaders() {
        Console.SetCursorPosition(OffsetLeft, OffsetTop);
        SetColors(ForeHeaderColor, BackHeaderColor);
        Console.Write(AlignText(" ", RowWidth, Cell.Alignments.Left));

        for(int r = 1; r < Console.WindowHeight - OffsetTop; r++) {
            Console.SetCursorPosition(OffsetLeft, OffsetTop + r);
            Console.Write(AlignText((r + StartRow).ToString(), RowWidth, Cell.Alignments.Left));
        }

        int c = 0;
        while(true) {
            int cc = SheetColumnToConsoleColumn(c);
            Console.SetCursorPosition(OffsetLeft + cc, OffsetTop);
            (string Text, bool Overflow) result = Trim(AlignText(GetColumnName(c + StartColumn), CellWidth, Cell.Alignments.Center), cc);
            Console.Write(result.Text);

            if(result.Overflow) break;
            c++;
        }
    }

    private (string Text, bool Overflow) Trim(string text, int cc) {
        if(cc + text.Length >= Console.WindowWidth - OffsetLeft) {
            return (text[..(Console.WindowWidth - OffsetLeft - cc)], true);
        }
        return (text, false);
    }

    private int SheetColumnToConsoleColumn(int c) {
        return c * CellWidth + RowWidth;
    }

    private string GetColumnName(int c) {
        StringBuilder sb = new();

        do {
            sb.Insert(0, (char)((c % ccCount) + 65));
            c /= ccCount;
            c--;
        } while(c >= 0);

        return sb.ToString();
    }

    internal string GetCellName(int col, int row) {
        return GetColumnName(col) + (row + 1).ToString();
    }

    internal string GetCellName(Cell c) {
         return GetCellName(c.Column, c.Row);
    }

    private (int Column, int Row) GetCellColRow(string name) {
        int c = 0;
        int k = 0;
        string rb = "";

        for(int i = name.Length - 1; i >= 0; i--) {
            if(char.IsDigit(name[i])) {
                rb = name[i] + rb;
            } else if(char.IsLetter(name[i])) {
                if(k == 0) {
                    c += (name[i] - 'A');
                } else {
                    c += (name[i] - '@') * (int)Math.Pow(10, k - 1) * ccCount;
                }
                k++;
            }
        }

        return (c, int.Parse(rb) - 1);
    }

    private string AlignText(string text, int width, Cell.Alignments alignment, int margin = 1) {
        string sm = new(' ', margin);

        return alignment switch {
            Cell.Alignments.Left => sm + text.PadRight(width - margin),
            Cell.Alignments.Center => (" ".PadLeft((int)Math.Ceiling((width - text.Length) / 2.0)) + text).PadRight(width),
            Cell.Alignments.Right => text.PadLeft(width - margin) + sm,
            _ => text,
        };
    }

    private void SetColors(ConsoleColor fore, ConsoleColor back) {
        Console.ForegroundColor = fore;
        Console.BackgroundColor = back;
    }
}