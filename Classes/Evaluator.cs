using NCalc;

internal class Evaluator {
    public const double Infinity = 10 ^ 6;
    public const double ToRad = Math.PI / 180;

    public delegate void CustomFunctionDel(string name, FunctionArgs args);

    private string formula = "";
    private readonly Dictionary<string, double> customParameters = [];
    private CustomFunctionDel? customFunction;
    private Expression? exp;

    private readonly Random rnd = new();
    private readonly Dictionary<int, string> strings = [];

    public CustomFunctionDel? CustomFunctionHandler {
        get { return customFunction; }
        set { customFunction = value; }
    }

    public Dictionary<string, object>? Variables { get => exp?.Parameters; }

    public Dictionary<string, double> CustomParameters { get => customParameters; }

    public Dictionary<int, string> Strings { get => strings; }

    public string Formula {
        get { return formula; }
        set {
            formula = value;
            if(formula == "") formula = "0";
            exp = new Expression(formula);

            exp.EvaluateFunction += (name, args) => {
                switch(name.ToUpper()) {
                    case "IIF":
                        if(args.Parameters.Length != 3) throw new ArgumentException("IIF function requires exactly 3 parameters");
                        args.Result = Convert.ToBoolean(args.Parameters[0].Evaluate()) ? args.Parameters[1].Evaluate() : args.Parameters[2].Evaluate();
                        break;
                    case "TORAD":
                        if(args.Parameters.Length != 1) throw new ArgumentException("TORAD function requires exactly 1 parameter");
                        args.Result = Convert.ToDouble(args.Parameters[0].Evaluate()) * ToRad;
                        break;
                    case "TODEG":
                        if(args.Parameters.Length != 1) throw new ArgumentException("TODEG function requires exactly 1 parameter");
                        args.Result = Convert.ToDouble(args.Parameters[0].Evaluate()) / ToRad;
                        break;
                    case "ABS":
                        if(args.Parameters.Length != 1) throw new ArgumentException("ABS function requires exactly 1 parameter");
                        args.Result = Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate()));
                        break;
                    case "RND":
                        if(args.Parameters.Length != 0) throw new ArgumentException("RND function requires exactly 0 parameters");
                        args.Result = rnd.NextDouble();
                        break;
                    case "MOD":
                        if(args.Parameters.Length != 2) throw new ArgumentException("MOD function requires exactly 2 parameters");
                        args.Result = Convert.ToDouble(args.Parameters[0].Evaluate()) % Convert.ToDouble(args.Parameters[1].Evaluate());
                        break;

                    case "SUM":
                        if(args.Parameters.Length < 2) throw new ArgumentException("SUM function requires 2 parameters or more");
                        args.Result = args.Parameters.Sum(p => Convert.ToDouble(p.Evaluate()));
                        args.Result = args.Parameters.Sum(p => Convert.ToDouble(p.Evaluate()));
                        break;
                    case "AVG":
                        if(args.Parameters.Length < 2) throw new ArgumentException("AVG function requires 2 parameters or more");
                        args.Result = args.Parameters.Average(p => Convert.ToDouble(p.Evaluate()));
                        break;

                    case "STR":
                        if(args.Parameters.Length != 1) throw new ArgumentException("STR function requires exactly 1 parameter");
                        args.Result = strings[Convert.ToInt32(args.Parameters[0].Evaluate())];
                        break;
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
                        if(customParameters.ContainsKey(name)) args.Result = customParameters[name];
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