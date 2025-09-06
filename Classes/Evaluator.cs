using NCalc;
using System.ComponentModel;

internal class Evaluator {
    public const double Infinity = 10 ^ 6;
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
        public string Name { get; set; } = "";
        public int MinParamCount { get; set; } = 0; // -1 for variable number of parameters
        public MathFunc Func { get; set; } = null;
        public string Description { get; set; } = "";

        public MathFuncDef() { }
    }

    delegate object MathFunc(FunctionArgs x);
    private readonly List<MathFuncDef> functions;

    public Evaluator() {
        functions = [
            new() { Name = "ABS",   MinParamCount = 1, Func = args => Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the absolute value of parameter" },
            new() { Name = "AVG",   MinParamCount = 2, Func = args => args.Parameters.Average(p => Convert.ToDouble(p.Evaluate())), Description = "Calculates the average (arithmetic mean) of parameters" },
            new() { Name = "COS",   MinParamCount = 1, Func = args => Math.Cos(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the cosine of parameter (in radians)" },
            new() { Name = "EXP",   MinParamCount = 1, Func = args => Math.Exp(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates e raised to the power of parameter" },
            new() { Name = "IIF",   MinParamCount = 3, Func = args => Convert.ToBoolean(args.Parameters[0].Evaluate()) ? args.Parameters[1].Evaluate() : args.Parameters[2].Evaluate(), Description = "If the first parameter is true, returns the second parameter, otherwise the third parameter" },
            new() { Name = "INT",   MinParamCount = 1, Func = args => Math.Floor(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the integer part of parameter" },
            new() { Name = "LN",    MinParamCount = 1, Func = args => Math.Log(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the natural logarithm of parameter" },
            new() { Name = "LOG10", MinParamCount = 1, Func = args => Math.Log10(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the base-10 logarithm of parameter" },
            new() { Name = "LOG2",  MinParamCount = 1, Func = args => Math.Log2(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the base-2 logarithm of parameter" },
            new() { Name = "MAX",   MinParamCount = 2, Func = args => args.Parameters.Max(p => Convert.ToDouble(p.Evaluate())), Description = "Calculates the maximum value of parameters" },
            new() { Name = "MIN",   MinParamCount = 2, Func = args => args.Parameters.Min(p => Convert.ToDouble(p.Evaluate())), Description = "Calculates the minimum value of parameters" },
            new() { Name = "MOD",   MinParamCount = 2, Func = args => Convert.ToDouble(args.Parameters[0].Evaluate()) % Convert.ToDouble(args.Parameters[1].Evaluate()), Description = "Calculates the modulus of two parameters" },
            new() { Name = "POW",   MinParamCount = 2, Func = args => Math.Pow(Convert.ToDouble(args.Parameters[0].Evaluate()), Convert.ToDouble(args.Parameters[1].Evaluate())), Description = "Calculates the power of the first parameter raised to the second parameter" },
            new() { Name = "RAND",  MinParamCount = 0, Func = args => new Random().NextDouble(), Description = "Generates a random number between 0 and 1" },
            new() { Name = "SIN",   MinParamCount = 1, Func = args => Math.Sin(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the sine of parameter (in radians)" },
            new() { Name = "SQRT",  MinParamCount = 1, Func = args => Math.Sqrt(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the square root of parameter" },
            new() { Name = "STD",   MinParamCount = 2, Func = args => {
                    double mean = args.Parameters.Average(p => Convert.ToDouble(p.Evaluate()));
                    double sumOfSquares = args.Parameters.Select(v => Math.Pow(Convert.ToDouble(v.Evaluate()) - mean, 2)).Sum();
                    return Math.Sqrt(sumOfSquares / args.Parameters.Length);
                },
                Description = "Calculates the standard deviation of parameters"
            },
            new() { Name = "STDS",  MinParamCount = 2, Func = args => {
                    double mean = args.Parameters.Average(p => Convert.ToDouble(p.Evaluate()));
                    double sumOfSquares = args.Parameters.Select(v => Math.Pow(Convert.ToDouble(v.Evaluate()) - mean, 2)).Sum();
                    return Math.Sqrt(sumOfSquares / (args.Parameters.Length - 1));
                },
                Description = "Calculates the sample standard deviation of parameters"
            },
            new() { Name = "STR",   MinParamCount = 1, Func = args =>  strings[Convert.ToInt32(args.Parameters[0].Evaluate())], Description = "Retrieves the string at the specified index" },
            new() { Name = "SUM",   MinParamCount = 2, Func = args => args.Parameters.Sum(p => Convert.ToDouble(p.Evaluate())), Description = "Calculates the sum of parameters" },
            new() { Name = "TAN",   MinParamCount = 1, Func = args => Math.Tan(Convert.ToDouble(args.Parameters[0].Evaluate())), Description = "Calculates the tangent of parameter (in radians)" },
            new() { Name = "TODEG", MinParamCount = 1, Func = args => Convert.ToDouble(args.Parameters[0].Evaluate()) / ToRad, Description = "Converts radians to degrees" },
            new() { Name = "TORAD", MinParamCount = 1, Func = args => Convert.ToDouble(args.Parameters[0].Evaluate()) * ToRad, Description = "Converts degrees to radians" }
        ];
    }

    public string Formula {
        get { return formula; }
        set {
            formula = value;
            if(formula == "") formula = "0";
            exp = new Expression(formula);

            exp.EvaluateFunction += (name, args) => {
                var func = functions.FirstOrDefault(f => f.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase));
                if(func != null) {
                    if(args.Parameters.Length < func.MinParamCount) throw new ArgumentException($"{func.Name} function requires at least {func.MinParamCount} parameter{(func.MinParamCount == 1 ? "" : "s")}");
                    args.Result = func.Func(args);
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

    private static double CalculateStandardDeviation(Expression[] parameters, bool sample) {
        double[] values = parameters.Select(p => Convert.ToDouble(p.Evaluate())).ToArray();
        double mean = values.Average();
        double sumOfSquares = values.Select(v => Math.Pow(v - mean, 2)).Sum();
        return Math.Sqrt(sumOfSquares / (values.Length - (sample ? 1 : 0)));
    }

    public (double Val, string? Str) Evaluate() {
        if(exp == null || formula == "0") return (0, null);

        var result = exp.Evaluate();
        if(result is string str) return (0, str);
        return (Convert.ToDouble(result), null);
    }
}