using System.Text;

internal class Sheet {
    public int StartColumn { get; set; } = 0;
    public int StartRow { get; set; } = 0;
    public int SelColumn { get; set; } = 0;
    public int SelRow { get; set; } = 0;
    public int SelFormulaColumn { get; set; } = 0;
    public int SelFormulaRow { get; set; } = 0;

    public int OffsetLeft { get; set; } = 0;
    public int OffsetTop { get; set; } = 3;

    public List<Cell> Cells { get; init; } = [];
    public int CellWidth { get; set; } = 15;
    public int RowWidth { get; set; } = 4;

    public ConsoleColor ForeCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackCellColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor ForeHeaderColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor BackHeaderColor { get; set; } = ConsoleColor.Cyan;
    public ConsoleColor BackHeaderSelColor { get; set; } = ConsoleColor.DarkCyan;
    public ConsoleColor ForeSelCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackSelCellColor { get; set; } = ConsoleColor.Blue;

    public string FileName { get; set; } = "sheet.csv";

    private readonly int ccCount = 'Z' - 'A' + 1; // 26
    private readonly string emptyCell;
    private string userInput = "";
    private int editCursorPosition = 0;

    enum Modes {
        Default,
        Edit,
        Formula,
        File
    }

    private Modes workingMode = Modes.Default;

    private readonly Dictionary<string, (string Key, string Action)[]> helpMessages = new() {
        { "default", new[] { ("Arrows", "Move"), ("Enter", "Edit"), ("Delete", "Delete Cell"), ("=", "Formula Mode"), ("'", "Label Mode"), ("\\", "File"), ("^Q", "Quit") } },
        { "edit", new[] { ("Enter", "Apply"), ("Esc", "Exit Edit Mode") } },
        { "formula", new[] { ("Arrows", "Select Cell"), ("Enter", "Add Cell to Formula"), ("Esc", "Exit Formula Mode") } },
        { "file", new[] { ("L", "Load Sheet"), ("S", "Save Sheet"), ("Esc", "Exit File Mode") } }
    };

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
        Cell? cell;

        while(true) {
            Render();

            Console.SetCursorPosition(0, 1);
            WriteLine(userInput);
            Console.SetCursorPosition(editCursorPosition, 1);

            ConsoleKeyInfo ck = Console.ReadKey(true);

            switch(workingMode) {
                case Modes.Default:
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

                        case ConsoleKey.Q:
                            if((ck.Modifiers & ConsoleModifiers.Control) == ConsoleModifiers.Control) {
                                return;
                            }
                            break;

                        case ConsoleKey.Backspace:
                            if(userInput.Length > 0)
                                userInput = userInput[..^1];
                            break;

                        case ConsoleKey.Enter:
                            cell = GetCell(SelColumn, SelRow);
                            if(cell != null) {
                                userInput = cell.ValueFormat;
                                editCursorPosition = userInput.Length;
                                workingMode = cell.Type == Cell.Types.Formula ? Modes.Formula : Modes.Edit;
                            }
                            break;

                        case ConsoleKey.Delete:
                            cell = GetCell(SelColumn, SelRow);
                            if(cell != null) {
                                Cells.Remove(cell);
                                string name = GetCellName(SelColumn, SelRow);
                                CascadeUpdate(name);
                            }
                            break;

                        case ConsoleKey.None:
                        case ConsoleKey.Oem5: // '\'
                            if(ck.KeyChar != '\\') break;
                            workingMode = Modes.File;
                            break;

                        default:
                            if(userInput.Length == 0 && ck.KeyChar == '=') {
                                SelFormulaColumn = SelColumn;
                                SelFormulaRow = SelRow;
                                workingMode = Modes.Formula;
                            } else {
                                if(!char.IsAsciiLetterOrDigit(ck.KeyChar) && ck.KeyChar != '\'') break;
                                workingMode = Modes.Edit;
                            }
                            if(userInput.Length < Console.WindowWidth - OffsetLeft - RowWidth)
                                userInput += ck.KeyChar;
                            editCursorPosition = userInput.Length;

                            break;
                    }
                    break;

                case Modes.Edit:
                case Modes.Formula:
                    switch(ck.Key) {
                        case ConsoleKey.UpArrow:
                            if(workingMode == Modes.Formula) {
                                if(SelFormulaRow > 0) SelFormulaRow--;
                            }
                            break;

                        case ConsoleKey.DownArrow:
                            if(workingMode == Modes.Formula)
                                SelFormulaRow++;
                            break;

                        case ConsoleKey.LeftArrow:
                            if(workingMode == Modes.Formula) {
                                if(SelFormulaColumn > 0) SelFormulaColumn--;
                            } else {
                                editCursorPosition = Math.Max(0, editCursorPosition - 1);
                            }
                            break;

                        case ConsoleKey.RightArrow:
                            if(workingMode == Modes.Formula) {
                                SelFormulaColumn++;
                            } else {
                                editCursorPosition = Math.Min(userInput.Length, editCursorPosition + 1);
                            }
                            break;

                        case ConsoleKey.Home:
                            editCursorPosition = 0;
                            break;

                        case ConsoleKey.End:
                            editCursorPosition = userInput.Length;
                            break;

                        case ConsoleKey.Enter:
                            if(workingMode == Modes.Formula) {
                                string name = GetCellName(SelFormulaColumn, SelFormulaRow);
                                userInput = userInput[0..editCursorPosition] + name + userInput[editCursorPosition..];
                                editCursorPosition += name.Length;
                            } else {
                                if(userInput.Length > 0) {
                                    cell = GetCell(SelColumn, SelRow);
                                    if(cell == null) {
                                        cell = new Cell(this, SelColumn, SelRow);
                                        Cells.Add(cell);
                                    }
                                    cell.Value = userInput;
                                    string name = GetCellName(SelColumn, SelRow);
                                    CascadeUpdate(name);
                                    userInput = "";
                                    editCursorPosition = 0;
                                    workingMode = Modes.Default;
                                }
                            }
                            break;

                        case ConsoleKey.Escape:
                            if(workingMode == Modes.Formula) {
                                cell = GetCell(SelColumn, SelRow);
                                if(cell == null) {
                                    cell = new Cell(this, SelColumn, SelRow);
                                    Cells.Add(cell);
                                }
                                cell.Value = userInput;
                            }
                            userInput = "";
                            editCursorPosition = 0;
                            workingMode = Modes.Default;
                            break;

                        case ConsoleKey.Backspace:
                            if(userInput.Length > 0) {
                                editCursorPosition--;
                                userInput = userInput[0..editCursorPosition] + userInput[(editCursorPosition + 1)..];
                            }
                            break;

                        case ConsoleKey.Delete:
                            if(editCursorPosition < userInput.Length) {
                                userInput = userInput[0..editCursorPosition] + userInput[(editCursorPosition + 1)..];
                            }
                            break;

                        default:
                            if(userInput.Length < Console.WindowWidth - OffsetLeft - RowWidth) {
                                userInput = userInput[0..editCursorPosition] + ck.KeyChar + userInput[editCursorPosition..];
                                editCursorPosition++;
                            }
                            break;
                    }
                    break;

                case Modes.File:
                    switch(ck.Key) {
                        case ConsoleKey.L:
                            Load();
                            workingMode = Modes.Default;
                            break;

                        case ConsoleKey.S:
                            Save();
                            workingMode = Modes.Default;
                            break;

                        case ConsoleKey.Escape:
                            workingMode = Modes.Default;
                            break;

                    }
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

    private void RenderHeaders() {
        SetColors(ConsoleColor.DarkGray, BackCellColor);

        RenderHelp(OffsetLeft, OffsetTop - 1, workingMode.ToString().ToUpper(), helpMessages[workingMode.ToString().ToLower()]);

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
                if((workingMode == Modes.Formula) && c == SelFormulaColumn && r == SelFormulaRow) {
                    SetColors(ConsoleColor.White, ConsoleColor.Red);
                } else {
                    SetColors(ForeCellColor, BackCellColor);
                }
            }

            Cell? cell = GetCell(c + StartColumn, r + StartRow);
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
            }
            ;
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

    internal (int Column, int Row) GetCellColRow(string name) {
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

    private static string AlignText(string text, int width, Cell.Alignments alignment, int margin = 1) {
        string sm = new(' ', margin);

        return alignment switch {
            Cell.Alignments.Left => sm + text.PadRight(width - margin),
            Cell.Alignments.Center => (" ".PadLeft((int)Math.Ceiling((width - text.Length) / 2.0)) + text).PadRight(width),
            Cell.Alignments.Right => text.PadLeft(width - margin) + sm,
            _ => text,
        };
    }

    private static void SetColors(ConsoleColor fore, ConsoleColor back) {
        Console.ForegroundColor = fore;
        Console.BackgroundColor = back;
    }

    public void Load() {
        if(File.Exists(FileName)) {
            Cells.Clear();
            string[] lines = File.ReadAllLines(FileName);
            for(int r = 0; r < lines.Length; r++) {
                string[] values = lines[r].Split(',');
                for(int c = 0; c < values.Length; c++) {
                    if(values[c].Trim() != "") {
                        Cell cell = new(this, c, r, values[c]);
                        Cells.Add(cell);
                    }
                }
            }
        }
    }

    public void Save() {
        if(FileName == "") return;
        StringBuilder sb = new();

        int maxColumn = Cells.Max(c => c.Column);
        int maxRow = Cells.Max(c => c.Row);

        for(int r = 0; r <= maxRow; r++) {
            for(int c = 0; c <= maxColumn; c++) {
                Cell? cell = GetCell(c, r);
                if(cell == null) {
                    sb.Append("");
                } else {
                    sb.Append(cell.ValueFormat);
                }
                if(c < maxColumn) sb.Append(',');
            }
            sb.AppendLine();
        }

        File.WriteAllText(FileName, sb.ToString());
    }
}
