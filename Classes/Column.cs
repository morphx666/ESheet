internal partial class Sheet {
    private const int DefaultColumnWidth = 15;

    public class Column {
        private int lastWidth = DefaultColumnWidth;
        private string emptyCell = AlignText(" ", DefaultColumnWidth, Cell.Alignments.Left);

        public int Index { get; set; }
        public int Width { get; set; } = DefaultColumnWidth;
        public string EmptyCell {
            get {
                if(lastWidth != Width) {
                    lastWidth = Width;
                    emptyCell = AlignText(" ", Width, Cell.Alignments.Left);
                }
                return emptyCell;
            }
        }

        public static int GetColumnWidth(Sheet sheet, int columnIndex) {
            var column = sheet.Columns.FirstOrDefault(c => c.Index == columnIndex);
            return column?.Width ?? DefaultColumnWidth;
        }

        public static string GetEmptyCell(Sheet sheet, int columnIndex) {
            var column = sheet.Columns.FirstOrDefault(c => c.Index == columnIndex);
            return column?.EmptyCell ?? AlignText(" ", DefaultColumnWidth, Cell.Alignments.Left);
        }
    }
}
