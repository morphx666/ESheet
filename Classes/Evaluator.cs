using NCalc;
using System.ComponentModel;

internal class Evaluator {
    public const double ToRad = Math.PI / 180;

    public delegate void CustomFunctionDel(string name, FunctionArgs args);

    private string formula = "";
    private readonly Dictionary<string, (double Val, string? Str)> customParameters = [];
    private CustomFunctionDel? customFunction;
    private Expression? exp;

    private readonly Random rnd = new();
    private readonly Dictionary<int, string> strings = [];

    public CustomFunctionDel? CustomFunctionHandler {
        get { return customFunction; }
        set { customFunction = value; }
    }

    public Dictionary<string, object>? Variables { get => exp?.Parameters; }

    public Dictionary<string, (double Val, string? Str)> CustomParameters { get => customParameters; }

    public Dictionary<int, string> Strings { get => strings; }

    class MathFuncDef {
        public string Name { get; init; }
        public int MinParamCount { get; init; }
        public MathFunc Function { get; init; }
        public string Description { get; init; }

        public MathFuncDef(string name, int minParamCount, MathFunc function, string description = "") {
            Name = name;
            MinParamCount = minParamCount;
            Function = function;
            Description = description;
        }
    }

    private delegate object MathFunc(FunctionArgs x);
    private readonly static List<MathFuncDef> functions = [];
    public static string[] FunctionNames {  get => [.. functions.Select(f => f.Name)]; }

    public Evaluator() {
        if(functions.Count > 0) return;
        functions.AddRange([
            new("ABS",      1, args => Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates the absolute value of parameter"),
            new("AVG",      2, args => args.Parameters.Average(p => Convert.ToDouble(p.Evaluate())),    "Calculates the average (arithmetic mean) of parameters"),
            new("COS",      1, args => Math.Cos(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates the cosine of parameter (in radians)"),
            new("EXP",      1, args => Math.Exp(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates e raised to the power of parameter"),
            new("IIF",      3, args => Convert.ToBoolean(args.Parameters[0].Evaluate())? args.Parameters[1].Evaluate() : args.Parameters[2].Evaluate(), "If the first parameter is true, returns the second parameter, otherwise the third parameter"),
            new("INT",      1, args => Math.Floor(Convert.ToDouble(args.Parameters[0].Evaluate())),     "Calculates the integer part of parameter"),
            new("LN",       1, args => Math.Log(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates the natural logarithm of parameter"),
            new("LOG10",    1, args => Math.Log10(Convert.ToDouble(args.Parameters[0].Evaluate())),     "Calculates the base-10 logarithm of parameter"),
            new("LOG2",     1, args => Math.Log2(Convert.ToDouble(args.Parameters[0].Evaluate())),      "Calculates the base-2 logarithm of parameter"),
            new("MAX",      2, args => args.Parameters.Max(p => Convert.ToDouble(p.Evaluate())),        "Calculates the maximum value of parameters"),
            new("MIN",      2, args => args.Parameters.Min(p => Convert.ToDouble(p.Evaluate())),        "Calculates the minimum value of parameters"),
            new("MOD",      2, args => Convert.ToDouble(args.Parameters[0].Evaluate()) % Convert.ToDouble(args.Parameters[1].Evaluate()), "Calculates the modulus of two parameters"),
            new("POW",      2, args => Math.Pow(Convert.ToDouble(args.Parameters[0].Evaluate()), Convert.ToDouble(args.Parameters[1].Evaluate())), "Calculates the power of the first parameter raised to the second parameter"),
            new("RAND",     0, args => new Random().NextDouble(),                                       "Generates a random number between 0 and 1"),
            new("SIN",      1, args => Math.Sin(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates the sine of parameter (in radians)"),
            new("SQRT",    1, args => Math.Sqrt(Convert.ToDouble(args.Parameters[0].Evaluate())),      "Calculates the square root of parameter"),
            new("STD",      2, args => {
                    double mean = args.Parameters.Average(p => Convert.ToDouble(p.Evaluate()));
                    double sumOfSquares = args.Parameters.Select(v => Math.Pow(Convert.ToDouble(v.Evaluate()) - mean, 2)).Sum();
                    return Math.Sqrt(sumOfSquares / args.Parameters.Length);
                },
                "Calculates the standard deviation of parameters"
            ),
            new("STDS",     2, args => {
                    double mean = args.Parameters.Average(p => Convert.ToDouble(p.Evaluate()));
                    double sumOfSquares = args.Parameters.Select(v => Math.Pow(Convert.ToDouble(v.Evaluate()) - mean, 2)).Sum();
                    return Math.Sqrt(sumOfSquares / (args.Parameters.Length - 1));
                },
                "Calculates the sample standard deviation of parameters"
            ),
            new("STR",      1, args => strings[Convert.ToInt32(args.Parameters[0].Evaluate())],         "Retrieves the string at the specified index"),
            new("SUM",      2, args => args.Parameters.Sum(p => Convert.ToDouble(p.Evaluate())),        "Calculates the sum of parameters"),
            new("TAN",      1, args => Math.Tan(Convert.ToDouble(args.Parameters[0].Evaluate())),       "Calculates the tangent of parameter (in radians)"),
            new("TODEG",    1, args => Convert.ToDouble(args.Parameters[0].Evaluate()) / ToRad,         "Converts radians to degrees"),
            new("TORAD",    1, args => Convert.ToDouble(args.Parameters[0].Evaluate()) * ToRad,         "Converts degrees to radians")
        ]);
    }

    public string Formula {
        get => formula;
        set {
            formula = value;
            if(formula == "") formula = "0";
            exp = new Expression(formula);

            exp.EvaluateFunction += (name, args) => {
                var func = functions.FirstOrDefault(f => f.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase));
                if(func != null) {
                    if(args.Parameters.Length < func.MinParamCount) throw new ArgumentException($"{func.Name} function requires at least {func.MinParamCount} parameter{(func.MinParamCount == 1 ? "" : "s")}");
                    args.Result = func.Function(args);
                    return;
                }
            };

            exp.EvaluateParameter += (name, args) => {
                switch(name.ToUpper()) {
                    case "PI":
                        args.Result = Math.PI;
                        break;
                    case "E":
                        args.Result = Math.E;
                        break;
                    case "C":
                        args.Result = 299_792_458; // Speed of light in m/s
                        break;
                    default:
                        if(customParameters.TryGetValue(name, out (double Val, string? Str) value)) args.Result = value.Val;
                        break;
                }
            };
        }
    }

    public (double Val, string? Str) Evaluate() {
        if(exp == null || formula == "0") return (0, null);

        var result = exp.Evaluate();
        if(result is string str) return (0, str);
        return (Convert.ToDouble(result), null);
    }
}