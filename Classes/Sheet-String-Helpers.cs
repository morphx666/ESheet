internal partial class Sheet {
    private static string AlignText(string text, int width, Cell.Alignments alignment, int margin = 1) {
        string sm = new(' ', margin);

        return alignment switch {
            Cell.Alignments.Left => sm + text.PadRight(width - margin),
            Cell.Alignments.Center => (" ".PadLeft((int)Math.Ceiling((width - text.Length) / 2.0)) + text).PadRight(width),
            Cell.Alignments.Right => text.PadLeft(width - margin) + sm,
            _ => text,
        };
    }

    private (string Text, bool Overflow) Trim(string text, int cc) {
        if(cc + text.Length >= Console.WindowWidth - OffsetLeft) {
            return (text[..(Console.WindowWidth - OffsetLeft - cc)], true);
        }
        return (text, false);
    }

    private string WorkingModeToKey() {
        string label = workingMode.ToString();
        string result = label[0].ToString();

        for(int i = 1; i < label.Length; i++) {
            if(char.IsUpper(label[i])) result += "|";
            result += workingMode.ToString()[i];
        }

        return result.ToLower();
    }

    private string WorkingModeToString() {
        string label = workingMode.ToString();
        string result = label[0].ToString();

        for(int i = 1; i < label.Length; i++) {
            if(char.IsUpper(label[i])) result += " ";
            result += workingMode.ToString()[i];
        }

        return result.ToUpper();
    }
}