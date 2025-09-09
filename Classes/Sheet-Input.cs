#pragma warning disable IDE0039

internal partial class Sheet {
    private void HandleInput() {
        Cell? cell = GetCell(SelColumn, SelRow);

        Func<char, bool> IsCtrlChar = c => c == '\'' || c == '=' || c == '\\';
        Func<bool> UserInputHasCtrlChar = () => userInput.Length > 0 && IsCtrlChar(userInput[0]);

        ConsoleKeyInfo ck = Console.ReadKey(true);

        bool isOpKey = isLinux
                        ? (ck.Modifiers & ConsoleModifiers.Alt) == ConsoleModifiers.Alt
                        : (ck.Modifiers & ConsoleModifiers.Control) == ConsoleModifiers.Control;
        bool isShiftKey = (ck.Modifiers & ConsoleModifiers.Shift) == ConsoleModifiers.Shift;

        // FIXME: There has to be a way to simplify this mess!
        switch(workingMode) {
            case Modes.Default:
                switch(ck.Key) {
                    case ConsoleKey.UpArrow:
                        if(SelRow > 0) {
                            SelRow--;
                            if(isShiftKey && cell?.Type == Cell.Types.Formula) {
                                Cell? newCell = GetCell(SelColumn, SelRow) ?? new(this, SelColumn, SelRow);
                                newCell.Value = cell.Value;
                                newCell.ShiftRefCells(0, -1);
                                Cells.Add(newCell);
                                FullRefresh();
                            }
                        }
                        break;

                    case ConsoleKey.DownArrow:
                        SelRow++;
                        if(isShiftKey && cell?.Type == Cell.Types.Formula) {
                            Cell? newCell = GetCell(SelColumn, SelRow) ?? new(this, SelColumn, SelRow);
                            newCell.Value = cell.Value;
                            newCell.ShiftRefCells(0, 1);
                            Cells.Add(newCell);
                            FullRefresh();
                        }
                        break;

                    case ConsoleKey.LeftArrow:
                        if(SelColumn > 0) {
                            SelColumn--;
                            if(isShiftKey && cell?.Type == Cell.Types.Formula) {
                                Cell? newCell = GetCell(SelColumn, SelRow) ?? new(this, SelColumn, SelRow);
                                newCell.Value = cell.Value;
                                newCell.ShiftRefCells(-1, 0);
                                Cells.Add(newCell);
                                FullRefresh();
                            }
                        }
                        break;

                    case ConsoleKey.RightArrow:
                        SelColumn++;
                        if(isShiftKey && cell?.Type == Cell.Types.Formula) {
                            Cell? newCell = GetCell(SelColumn, SelRow) ?? new(this, SelColumn, SelRow);
                            newCell.Value = cell.Value;
                            newCell.ShiftRefCells(1, 0);
                            Cells.Add(newCell);
                            FullRefresh();
                        }
                        break;

                    case ConsoleKey.PageUp:
                        SelRow = Math.Max(0, SelRow - (Console.WindowHeight - OffsetTop - 1));
                        while((SelRow - StartRow) < 0) StartRow--; // TODO: Move this to a separate method
                        break;

                    case ConsoleKey.PageDown:
                        SelRow += Console.WindowHeight - OffsetTop - 1;
                        while((SelRow - StartRow) >= Console.WindowHeight - OffsetTop - 1) StartRow++; // TODO: Move this to a separate method
                        break;

                    case ConsoleKey.Backspace:
                        if(userInput.Length > 0)
                            userInput = userInput[..^1];
                        break;

                    case ConsoleKey.Enter:
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
                        while(SheetColumnToConsoleColumn(SelColumn - StartColumn) >= Console.WindowWidth - OffsetLeft) StartColumn++; // TODO: Move this to a separate method
                        SelRow = Cells.Count == 0 ? 0 : Cells.Max(c => c.Row);
                        while((SelRow - StartRow) >= Console.WindowHeight - OffsetTop - 1) StartRow++; // TODO: Move this to a separate method
                        break;

                    default:
                        switch(ck.Key) {
                            case ConsoleKey.Q:
                                if(isOpKey) {
                                    Console.ResetColor();
                                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                                    Console.WriteLine();
                                    Environment.Exit(0);
                                }
                                break;

                            case ConsoleKey.K:
                                if(isOpKey) {
                                    workingMode = Modes.Sheet;
                                    return;
                                }
                                break;
                        }

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
                        if(IsCtrlChar(ck.KeyChar)
                            || char.IsAsciiLetterOrDigit(ck.KeyChar)
                            || ck.KeyChar == '-'
                            || ck.KeyChar == '+'
                            || ck.KeyChar == '"') {
                            if(userInput.Length < Console.WindowWidth - OffsetLeft - RowHeaderWidth)
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
                        if(workingMode == Modes.Formula && isOpKey) {
                            if(SelFormulaRow > 0) SelFormulaRow--;
                        }
                        break;

                    case ConsoleKey.DownArrow:
                        if(workingMode == Modes.Formula && isOpKey) {
                            SelFormulaRow++;
                        }
                        break;

                    case ConsoleKey.LeftArrow:
                        if(workingMode == Modes.Formula && isOpKey) {
                            if(SelFormulaColumn > 0) SelFormulaColumn--;
                        } else {
                            editCursorPosition = Math.Max(workingMode == Modes.Formula ? 1 : 0, editCursorPosition - 1);
                        }
                        break;

                    case ConsoleKey.RightArrow:
                        if(workingMode == Modes.Formula && isOpKey) {
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
                        if(workingMode == Modes.Formula && isOpKey) {
                            string name = GetCellName(SelFormulaColumn, SelFormulaRow);
                            userInput = userInput[0..editCursorPosition] + name + userInput[editCursorPosition..];
                            editCursorPosition += name.Length;
                        } else if(userInput.Length > 0) {
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
                        if(userInput.Length < Console.WindowWidth - OffsetLeft - RowHeaderWidth) {
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
                            ResetSheet();
                            FileName = "esheet.csv";
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
                                LoadFile(userInput);
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
                        if((workingMode == Modes.FileLoad || workingMode == Modes.FileSave) && userInput.Length < Console.WindowWidth - OffsetLeft - RowHeaderWidth) {
                            userInput = userInput[0..editCursorPosition] + ck.KeyChar + userInput[editCursorPosition..];
                            editCursorPosition++;
                        }
                        break;

                }
                break;

            case Modes.Sheet:
            case Modes.SheetRow:
            case Modes.SheetColumn:
                switch(ck.Key) {
                    case ConsoleKey.LeftArrow:
                        if(isOpKey) {
                            InsertColumn(SelColumn + StartColumn, -1);
                        } else {
                            if(SelColumn > 0) SelColumn--;
                        }
                        break;

                    case ConsoleKey.RightArrow:
                        if(isOpKey) {
                            InsertColumn(SelColumn + StartColumn, 1);
                        } else {
                            SelColumn++;
                        }
                        break;

                    case ConsoleKey.UpArrow:
                        if(isOpKey) {
                            InsertRow(SelRow + StartRow, -1);
                        } else {
                            if(SelRow > 0) SelRow--;
                        }
                        break;

                    case ConsoleKey.DownArrow:
                        if(isOpKey) {
                            InsertRow(SelRow + StartRow, 1);
                        } else {
                            SelRow++;
                        }
                        break;

                    case ConsoleKey.Delete:
                        if(workingMode == Modes.SheetColumn) {
                            DeleteColumn(SelColumn + StartColumn);
                        } else {
                            DeleteRow(SelRow + StartRow);
                        }
                        break;

                    case ConsoleKey.OemPlus:
                        if(isOpKey) SetColumnWidth(SelColumn + StartColumn, 1);
                        break;

                    case ConsoleKey.OemMinus:
                        if(isOpKey) SetColumnWidth(SelColumn + StartColumn, -1);
                        break;

                    case ConsoleKey.C:
                        workingMode = Modes.SheetColumn;
                        break;

                    case ConsoleKey.R:
                        workingMode = Modes.SheetRow;
                        break;

                    case ConsoleKey.Escape:
                        workingMode = Modes.Default;
                        userInput = "";
                        editCursorPosition = 0;
                        break;
                }
                break;

        }
    }
}