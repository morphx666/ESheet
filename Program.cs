internal class Program {
    private readonly static Sheet sheet = new();

    private static void Main(string[] args) {
        Console.Clear();

        sheet.RenderPrecision = 2; // Display 2 decimal places for numeric values
        sheet.Run();

        Console.ResetColor();
        Console.Clear();
        Console.CursorVisible = true;
    }
}