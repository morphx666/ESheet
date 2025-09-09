internal class Program {
    private readonly static Sheet sheet = new();

    private static void Main(string[] args) {
        Console.Clear();

        sheet.Run();

        Console.ResetColor();
        Console.Clear();
        Console.CursorVisible = true;
    }
}