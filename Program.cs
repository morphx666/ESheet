using ESheet.Classes;
using System.Diagnostics;

internal class Program {
    private static Sheet sheet = new();

    private static void Main(string[] args) {
        Console.Clear();

        sheet.Run();

        Console.ResetColor();
        Console.Clear();
        Console.CursorVisible = true;
    }
}