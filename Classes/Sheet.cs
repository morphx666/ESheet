#pragma warning disable IDE0039

using System.Text;

internal partial class Sheet {
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

    public int RenderPrecision { get; set; } = 2;

    public ConsoleColor ForeCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackCellColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor ForeHeaderColor { get; set; } = ConsoleColor.Black;
    public ConsoleColor BackHeaderColor { get; set; } = ConsoleColor.Cyan;
    public ConsoleColor BackHeaderSelColor { get; set; } = ConsoleColor.DarkCyan;
    public ConsoleColor ForeSelCellColor { get; set; } = ConsoleColor.White;
    public ConsoleColor BackSelCellColor { get; set; } = ConsoleColor.Blue;

    private string fileName = "esheet.csv";
    public string FileName {
        get => fileName;
        set {
            fileName = value;
            Console.Title = $"ESheet - {fileName}";
        }
    }

    private readonly int ccCount = 'Z' - 'A' + 1; // 26
    private readonly string emptyCell;
    private string userInput = "";
    private int editCursorPosition = 0;

    enum Modes {
        Default,
        Edit,
        Formula,
        File,
        FileLoad,
        FileSave
    }

    private Modes workingMode = Modes.Default;

    private readonly Dictionary<string, (string Key, string Action)[]> helpMessages = new() {
        { "default", new[] { ("Arrows", "Move"), ("Enter", "Edit"), ("Delete", "Delete"), ("=", "Formula"), ("'", "Label"), ("\\", "File"), ("^Q", "Quit") } },
        { "edit", new[] { ("Enter", "Apply"), ("Esc", "Exit Edit Mode") } },
        { "formula", new[] { ("^Arrows", "Select Cell"), ("Enter", "Apply"), ("^Enter", "Add Cell to Formula"), ("Esc", "Exit Formula Mode") } },
        { "file", new[] { ("N", "New Sheet"), ("L", "Load Sheet"), ("S", "Save Sheet"), ("Esc", "Exit File Mode") } },
        { "file|load", new[] { ("Enter", "Load"), ("Esc", "Cancel Load") } },
        { "file|save", new[] { ("Enter", "Save"), ("Esc", "Cancel Save") } }
    };

    public Sheet() {
        Console.Title = $"ESheet";
        emptyCell = AlignText(" ", CellWidth, Cell.Alignments.Left);
    }

    public void Run() {
        Cell? cell;

        Func<char, bool> IsCtrlChar = c => {
            return c == '\'' || c == '=' || c == '\\';
        };

        Func<bool> UserInputHasCtrlChar = () => {
            return userInput.Length > 0 && IsCtrlChar(userInput[0]);
        };

        while(true) {
            Render();

            Console.SetCursorPosition(0, 1);
            WriteLine(userInput);
            Console.SetCursorPosition(editCursorPosition, 1);

            ConsoleKeyInfo ck = Console.ReadKey(true);

            // FIXME: There has to be a way to simplify this mess!
            bool isCtrl = (ck.Modifiers & ConsoleModifiers.Control) == ConsoleModifiers.Control;
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
                            if(isCtrl) return;
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
                                if(workingMode == Modes.Formula) {
                                    SelFormulaColumn = SelColumn;
                                    SelFormulaRow = SelRow;
                                }
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

                        case ConsoleKey.Home:
                            SelColumn = 0;
                            StartColumn = 0;
                            SelRow = 0;
                            StartRow = 0;
                            break;

                        case ConsoleKey.End:
                            SelColumn = Cells.Count == 0 ? 0 : Cells.Max(c => c.Column);
                            while(SheetColumnToConsoleColumn(SelColumn - StartColumn) >= Console.WindowWidth - OffsetLeft) StartColumn++;
                            SelRow = Cells.Count == 0 ? 0 : Cells.Max(c => c.Row);
                            while((SelRow - StartRow) >= Console.WindowHeight - OffsetTop - 1) StartRow++;
                            break;

                        default:
                            if(userInput.Length == 0) {
                                switch(ck.KeyChar) {
                                    case '=':
                                        SelFormulaColumn = SelColumn;
                                        SelFormulaRow = SelRow;
                                        workingMode = Modes.Formula;
                                        break;
                                    case '\\':
                                        workingMode = Modes.File;
                                        break;
                                    case '\'':
                                    default:
                                        workingMode = Modes.Edit;
                                        break;
                                }
                            }
                            if(IsCtrlChar(ck.KeyChar) || char.IsAsciiLetterOrDigit(ck.KeyChar)) {
                                if(userInput.Length < Console.WindowWidth - OffsetLeft - RowWidth)
                                    userInput += ck.KeyChar;
                                editCursorPosition = userInput.Length;
                            }
                            break;
                    }
                    break;

                case Modes.Edit:
                case Modes.Formula:
                    switch(ck.Key) {
                        case ConsoleKey.UpArrow:
                            if(workingMode == Modes.Formula && isCtrl) {
                                if(SelFormulaRow > 0) SelFormulaRow--;
                            }
                            break;

                        case ConsoleKey.DownArrow:
                            if(workingMode == Modes.Formula && isCtrl) {
                                SelFormulaRow++;
                            }
                            break;

                        case ConsoleKey.LeftArrow:
                            if(workingMode == Modes.Formula && isCtrl) {
                                if(SelFormulaColumn > 0) SelFormulaColumn--;
                            } else {
                                editCursorPosition = Math.Max(workingMode == Modes.Formula ? 1 : 0, editCursorPosition - 1);
                            }
                            break;

                        case ConsoleKey.RightArrow:
                            if(workingMode == Modes.Formula && isCtrl) {
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
                            if(workingMode == Modes.Formula && isCtrl) {
                                string name = GetCellName(SelFormulaColumn, SelFormulaRow);
                                userInput = userInput[0..editCursorPosition] + name + userInput[editCursorPosition..];
                                editCursorPosition += name.Length;
                            } else if(userInput.Length > 0) {
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
                            break;

                        case ConsoleKey.Escape:
                            userInput = "";
                            editCursorPosition = 0;
                            workingMode = Modes.Default;
                            break;

                        case ConsoleKey.Backspace:
                            if(userInput.Length > (UserInputHasCtrlChar() ? 1 : 0) && editCursorPosition > 0) {
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
                case Modes.FileLoad:
                case Modes.FileSave:
                    switch(ck.Key) {
                        case ConsoleKey.LeftArrow:
                            if(workingMode == Modes.FileLoad || workingMode == Modes.FileSave) {
                                editCursorPosition = Math.Max(0, editCursorPosition - 1);
                            }
                            break;

                        case ConsoleKey.RightArrow:
                            if(workingMode == Modes.FileLoad || workingMode == Modes.FileSave) {
                                editCursorPosition = Math.Min(userInput.Length, editCursorPosition + 1);
                            }
                            break;

                        case ConsoleKey.Home:
                            editCursorPosition = 0;
                            break;

                        case ConsoleKey.End:
                            editCursorPosition = userInput.Length;
                            break;

                        case ConsoleKey.Backspace:
                            if(userInput.Length > (UserInputHasCtrlChar() ? 1 : 0) && editCursorPosition > 0) {
                                editCursorPosition--;
                                userInput = userInput[0..editCursorPosition] + userInput[(editCursorPosition + 1)..];
                            }
                            break;

                        case ConsoleKey.Delete:
                            if(editCursorPosition < userInput.Length) {
                                userInput = userInput[0..editCursorPosition] + userInput[(editCursorPosition + 1)..];
                            }
                            break;

                        case ConsoleKey.N:
                            if(workingMode == Modes.File) {
                                Cells.Clear();
                                FileName = "esheet.csv";
                                userInput = "";
                                editCursorPosition = 0;
                                workingMode = Modes.Default;
                            } else goto handleFileModeKeyStroke;
                            break;

                        case ConsoleKey.L:
                            if(workingMode == Modes.File) {
                                workingMode = Modes.FileLoad;
                                userInput = FileName;
                                editCursorPosition = userInput.Length;
                            } else goto handleFileModeKeyStroke;
                            break;

                        case ConsoleKey.S:
                            if(workingMode == Modes.File) {
                                workingMode = Modes.FileSave;
                                userInput = FileName;
                                editCursorPosition = userInput.Length;
                            } else goto handleFileModeKeyStroke;
                            break;

                        case ConsoleKey.Enter:
                            switch(workingMode) {
                                case Modes.FileLoad:
                                    if(LoadFile(userInput)) {
                                        workingMode = Modes.Default;
                                        userInput = "";
                                        editCursorPosition = 0;
                                    }
                                    break;

                                case Modes.FileSave:
                                    if(SaveFile(userInput)) {
                                        workingMode = Modes.Default;
                                        userInput = "";
                                        editCursorPosition = 0;
                                    }
                                    break;
                            }
                            break;

                        case ConsoleKey.Escape:
                            workingMode = Modes.Default;
                            userInput = "";
                            editCursorPosition = 0;
                            break;

                        default:
handleFileModeKeyStroke:
                            if((workingMode == Modes.FileLoad || workingMode == Modes.FileSave) && userInput.Length < Console.WindowWidth - OffsetLeft - RowWidth) {
                                userInput = userInput[0..editCursorPosition] + ck.KeyChar + userInput[editCursorPosition..];
                                editCursorPosition++;
                            }
                            break;

                    }
                    break;
            }
        }
    }

    private void CascadeUpdate(string name) {
        List<string> cellsToUpdate = [];
        foreach(Cell c in Cells.ToArray()) { // Silly way to create a copy of the Cells collection and avoid "Collection was modified" exception
            if(c.Type == Cell.Types.Formula && c.ExpandRanges(c.Value).Contains(name)) {
                c.Refresh();
                string cellToUpdate = GetCellName(c);
                cellsToUpdate.Add(cellToUpdate);
            }
        }

        cellsToUpdate.ForEach(c => CascadeUpdate(c));
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
        try {
            (int col, int row) = GetCellColRow(name);
            return GetCell(col, row);
        } catch {
            return null;
        }
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

        if(int.TryParse(rb, out int v)) {
            return (c, v - 1);
        } else {
            throw new ArgumentException("Invalid cell name", nameof(name));
        }
    }

    internal (bool IsValid, int Column, int Row) IsCellNameValid(string name) {
        try {
            (int Column, int Row) = GetCellColRow(name);
            return (true, Column, Row);
        } catch {
            return (false, -1, -1);
        }
    }

    private void FullRefresh() {
        Cells.ForEach(c => {
            if(c.Type == Cell.Types.Formula) c.Refresh();
        });
    }

    public bool LoadFile(string fileName) {
        if(File.Exists(fileName)) {
            Cells.Clear();
            string[] lines = File.ReadAllLines(fileName);
            for(int r = 0; r < lines.Length; r++) {
                string[] values = lines[r].Split(',');
                for(int c = 0; c < values.Length; c++) {
                    if(values[c].Trim() != "") {
                        Cell cell = new(this, c, r, values[c]);
                        Cells.Add(cell);
                    }
                }
            }

            FullRefresh();

            FileName = fileName;
            return true;
        }
        return false;
    }

    public bool SaveFile(string fileName) {
        if(fileName == "") return false;
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

        File.WriteAllText(fileName, sb.ToString());
        FileName = fileName;

        return true;
    }
}
