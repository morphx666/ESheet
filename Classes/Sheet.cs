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
    public List<Column> Columns { get; init; } = [];
    public int RowHeaderWidth { get; set; } = 4;

    public int RenderPrecision { get; set; } = 2;

    private string fileName = "esheet.csv";
    public string FileName {
        get => fileName;
        set {
            fileName = value;
            Console.Title = $"ESheet - {fileName}";
        }
    }

    private readonly int ccCount = 'Z' - 'A' + 1; // 26
    private string userInput = "";
    private int editCursorPosition = 0;
    private string emptyLine = "";

    private static readonly bool isLinux = Environment.OSVersion.Platform == PlatformID.Unix;

    enum Modes {
        Invalid,

        Default,
        Edit,

        Formula,

        File,
        FileLoad,
        FileSave,

        Sheet,
        SheetColumn,
        SheetRow,
    }

    private Modes workingMode = Modes.Default;
    private Modes lastWorkingMode = Modes.Invalid;
    private readonly Dictionary<int, string> columnsNameCache = [];

    private readonly Dictionary<string, (string Key, string Action)[]> helpMessages = new() {
        { "default", new[] { ("←↑↓→", "Move"), ("Enter", "Edit"), ("Delete", "Delete"), ("=", "Formula"), ("'", "Label"), ("^K", "Sheet"), ("\\", "File"), ("^Q", "Quit") } },
        { "edit", new[] { ("Enter", "Apply"), ("Esc", "Exit Edit Mode") } },
        { "formula", new[] { ("^←↑↓→", "Select Cell"), ("Enter", "Apply"), ("^Enter", "Add Cell to Formula"), ("Esc", "Exit Formula Mode") } },
        { "file", new[] { ("N", "New Sheet"), ("L", "Load Sheet"), ("S", "Save Sheet"), ("Esc", "Exit File Mode") } },
        { "file|load", new[] { ("Enter", "Load"), ("Esc", "Cancel Load") } },
        { "file|save", new[] { ("Enter", "Save"), ("Esc", "Cancel Save") } },
        { "sheet", new[] { ("C", "Columns"), ("R", "Rows"), ("Esc", "Exit Sheet Mode") } },
        { "sheet|column", new[] { ("^←", "Insert Left"), ("^→", "Insert Right"), ("Delete", "Delete"), ("^+", "Wider"), ("^-", "Narrower"), ("Esc", "Exit Column Mode") } },
        { "sheet|row", new[] { ("^↑", "Insert Above"), ("^↓", "Insert Below"), ("Delete", "Delete"), ("Esc", "Exit Row Mode") } },
    };

    public Sheet() {
        Console.Title = "ESheet";
        Console.ResetColor();
    }

    public void Run() {
        while(true) {
MainLoop:
            Render();

            Console.SetCursorPosition(0, 1);
            WriteLine(userInput);
            Console.SetCursorPosition(editCursorPosition, 1);

            int w = Console.WindowWidth;
            int h = Console.WindowHeight;
            while(!Console.KeyAvailable) {
                if(w != Console.WindowWidth || h != Console.WindowHeight) {
                    Console.Clear();
                    lastWorkingMode = Modes.Invalid; // Force Help re-render
                    Thread.Sleep(250); // Wait for the console to stabilize
                    emptyLine = new(' ', Console.WindowWidth - 1);
                    goto MainLoop;
                }
                Thread.Sleep(60);
            }

            HandleInput();
        }
    }

    static int maxRecursion;
    private void CascadeUpdate(string name, bool init = true) {
        if(init) {
            maxRecursion = Cells.Count * 2;
        } else {
            maxRecursion--;
            if(maxRecursion == 0) {
                maxRecursion = -1;
                GetCell(name)?.SetError("Circular reference");
                return;
            }
        }

        List<string> cellsToUpdate = [];
        foreach(Cell c in Cells.ToArray()) { // Silly way to create a copy of the Cells collection and avoid "Collection was modified" exception
            if(!c.HasError && c.Type == Cell.Types.Formula && c.ExpandRanges(c.Value).Contains(name)) {
                c.Refresh();
                string cellToUpdate = GetCellName(c);
                cellsToUpdate.Add(cellToUpdate);
            }
        }

        cellsToUpdate.ForEach(c => CascadeUpdate(c, false));
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
        int cc = 0;
        for(int i = 0; i < c; i++) {
            cc += GetColumnWidth(i);
        }
        return cc + RowHeaderWidth;
    }

    private int GetColumnWidth(int c) {
        Column? column = Columns.FirstOrDefault(col => col.Index == c);
        return column?.Width ?? DefaultColumnWidth;
    }

    private string GetColumnName(int c) {
        if(columnsNameCache.TryGetValue(c, out string? value)) return value;

        int col = c;
        StringBuilder sb = new();

        do {
            sb.Insert(0, (char)((c % ccCount) + 65));
            c /= ccCount;
            c--;
        } while(c >= 0);

        columnsNameCache[col] = sb.ToString();
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
            if(Evaluator.FunctionNames.Contains(name.ToUpper())) return (false, -1, -1);
            (int Column, int Row) = GetCellColRow(name);
            return (true, Column, Row);
        } catch {
            return (false, -1, -1);
        }
    }

    internal void FullRefresh() {
        List<Cell> emptyCells = Cells.Select(c => c).Where(c => c.Value == "").ToList();
        while(emptyCells.Count > 0) {
            Cells.Remove(emptyCells[0]);
            emptyCells.RemoveAt(0);
        }

        Cells.ForEach(c => {
            if(c.Type == Cell.Types.Formula) c.Refresh();
        });
    }

    private void InsertColumn(int col, int direction) {
        Cells.ForEach(c => {
            switch(direction) {
                case -1:
                    if(c.Column >= col) {
                        int offset = 0;
                        var cellsInFormula = c.GetReferencedCells();
                        foreach(var cell in cellsInFormula) {
                            if(cell.Column >= col) {
                                string newName = GetCellName(cell.Column + 1, cell.Row);
                                c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                                offset += newName.Length - cell.Name.Length;
                            }
                        }
                        c.SetColRow(c.Column + 1, c.Row);
                    }
                    break;
                case 1:
                    if(c.Column >= col + 1) {
                        int offset = 0;
                        var cellsInFormula = c.GetReferencedCells();
                        foreach(var cell in cellsInFormula) {
                            if(cell.Column >= col + 1) {
                                string newName = GetCellName(cell.Column + 1, cell.Row);
                                c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                                offset += newName.Length - cell.Name.Length;
                            }
                        }
                        c.SetColRow(c.Column + 1, c.Row);
                    }
                    break;
            }
        });

        FullRefresh();
    }

    private void DeleteColumn(int col) {
        Cells.RemoveAll(c => c.Column == col);
        Cells.ForEach(c => {
            if(c.Column > col) {
                int offset = 0;
                var cellsInFormula = c.GetReferencedCells();
                foreach(var cell in cellsInFormula) {
                    if(cell.Column > col) {
                        string newName = GetCellName(cell.Column - 1, cell.Row);
                        c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                        offset += newName.Length - cell.Name.Length;
                    }
                }
                c.SetColRow(c.Column - 1, c.Row);
            }
        });

        FullRefresh();
    }

    private void DeleteRow(int row) {
        Cells.RemoveAll(c => c.Row == row);
        Cells.ForEach(c => {
            if(c.Row > row) {
                int offset = 0;
                var cellsInFormula = c.GetReferencedCells();
                foreach(var cell in cellsInFormula) {
                    if(cell.Row > row) {
                        string newName = GetCellName(cell.Column, cell.Row - 1);
                        c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                        offset += newName.Length - cell.Name.Length;
                    }
                }
                c.SetColRow(c.Column, c.Row - 1);
            }
        });

        FullRefresh();
    }

    private void InsertRow(int row, int direction) {
        Cells.ForEach(c => {
            switch(direction) {
                case -1:
                    if(c.Row >= row) {
                        int offset = 0;
                        var cellsInFormula = c.GetReferencedCells();
                        foreach(var cell in cellsInFormula) {
                            if(cell.Row >= row) {
                                string newName = GetCellName(cell.Column, cell.Row + 1);
                                c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                                offset += newName.Length - cell.Name.Length;
                            }
                        }
                        c.SetColRow(c.Column, c.Row + 1);
                    }
                    break;
                case 1:
                    if(c.Row >= row + 1) {
                        int offset = 0;
                        var cellsInFormula = c.GetReferencedCells();
                        foreach(var cell in cellsInFormula) {
                            if(cell.Row >= row + 1) {
                                string newName = GetCellName(cell.Column, cell.Row + 1);
                                c.SetValueFast(c.Value[0..(cell.Pos + offset)] + newName + c.Value[(cell.Pos + cell.Name.Length + offset)..]);
                                offset += newName.Length - cell.Name.Length;
                            }
                        }
                        c.SetColRow(c.Column, c.Row + 1);
                    }
                    break;
            }
        });

        FullRefresh();
    }

    private void SetColumnWidth(int col, int delta) {
        Column? column = Columns.FirstOrDefault(c => c.Index == col);
        if(column == null) {
            column = new() { Index = col };
            Columns.Add(column);
        }
        if(column.Width + delta > 0) {
            column.Width = Math.Max(1, column.Width + delta);
        }
    }

    private void ResetSheet() {
        Cells.Clear();
        Columns.Clear();

        userInput = "";
        editCursorPosition = 0;
        workingMode = Modes.Default;

        //StartColumn = 0;
        //SelColumn = 0;
        //StartRow = 0;
        //SelRow = 0;
    }

    public bool LoadFile(string fileName) {
        if(File.Exists(fileName)) {
            ResetSheet();

            string[] lines = File.ReadAllLines(fileName);
            for(int r = 0; r < lines.Length; r++) {
                string[] values = lines[r].Split('\t');
                if(values[0] == "#COLUMN") {
                    int colIndex = int.Parse(values[1]);
                    int colWidth = int.Parse(values[2]);
                    Column column = new() { Index = colIndex, Width = colWidth };
                    Columns.Add(column);
                } else {
                    for(int c = 0; c < values.Length; c++) {
                        if(values[c].Trim() != "") {
                            Cell? cell = GetCell(c, r);
                            if(cell is null) {
                                cell = new(this, c, r, values[c]);
                            } else {
                                cell.Value = values[c];
                            }
                            Cells.Add(cell);
                        }
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

        int maxRow = Cells.Max(c => c.Row);

        for(int r = 0; r <= maxRow; r++) {
            var cellsInRow = Cells.Where(c => c.Row == r);
            if(cellsInRow.Any()) {
                int maxColumn = cellsInRow.Max(c => c.Column);

                for(int c = 0; c <= maxColumn; c++) {
                    Cell? cell = GetCell(c, r);
                    if(cell == null) {
                        sb.Append("");
                    } else {
                        sb.Append(cell.ValueFormat);
                    }
                    if(c < maxColumn) sb.Append('\t');
                }
            }
            sb.AppendLine();
        }

        foreach(Column col in Columns.OrderBy(c => c.Index)) {
            if(col.Width != DefaultColumnWidth) {
                sb.AppendLine($"#COLUMN\t{col.Index}\t{col.Width}");
            }
        }

        File.WriteAllText(fileName, sb.ToString());
        FileName = fileName;

        return true;
    }
}