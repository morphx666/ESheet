using NCalc;
using System.ComponentModel;
using System.Diagnostics;

namespace ESheet.Classes {
    internal class Evaluator {
        public const double Infinity = 10 ^ 6;
        public const double ToRad = Math.PI / 180;

        public delegate void CustomFunctionDel(string name, FunctionArgs args);

        private string formula = "";
        private readonly Dictionary<string, double> customParameters = [];
        private CustomFunctionDel? customFunction;

        private Expression? exp;
        private readonly Random rnd = new();

        public CustomFunctionDel? CustomFunctionHandler {
            get { return customFunction; }
            set { customFunction = value; }
        }

        public string Formula {
            get { return formula; }
            set {
                formula = value;
                if(formula == "") formula = "0";
                exp = new Expression(formula);

                exp.EvaluateFunction += (name, args) => {
                    switch(name.ToUpper()) {
                        case "IIF":
                            args.Result = Convert.ToBoolean(args.Parameters[0].Evaluate()) ? args.Parameters[1].Evaluate() : args.Parameters[2].Evaluate();
                            break;
                        case "TORAD":
                            args.Result = Convert.ToDouble(args.Parameters[0].Evaluate()) * ToRad;
                            break;
                        case "ABS":
                            args.Result = Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate()));
                            break;
                        case "RND":
                            args.Result = rnd.NextDouble() - 0.5;
                            break;
                        case "MOD":
                            args.Result = Convert.ToDouble(args.Parameters[0].Evaluate()) % Convert.ToDouble(args.Parameters[1].Evaluate());
                            break;

                        case "SUM":
                            args.Result = args.Parameters.Sum(p => (double)p.Evaluate());
                            break;
                        case "AVG":
                            args.Result = args.Parameters.Average(p => (double)p.Evaluate());
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

        public Dictionary<string, object>? Variables {
            get { return exp?.Parameters; }
        }

        public Dictionary<string, double> CustomParameters {
            get { return customParameters; }
        }

        public double Evaluate() {
            return exp == null || formula == "0" ? 0 : Convert.ToDouble(exp.Evaluate());
        }
    }
}
