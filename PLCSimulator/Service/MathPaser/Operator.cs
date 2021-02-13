using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLCTools.Service.MathPaser
{
    public class Operator
    {
        public Operator(OperatorType type, string value)
        {
            this.Type = type;
            this.Value = value;
        }

        /// <summary>
        /// Operator type
        /// </summary>
        public OperatorType Type { get; set; }

        /// <summary>
        /// Operater value
        /// </summary>
        public string Value { get; set; }


        /// <summary>
        // For > or &lt; operators, determine if the actual is >=,&lt;&gt;, &lt;=, and adjust the current operator position
        /// </summary>
        // <param name="currentOpt">Current Operator</param>
        // <param name="currentExp">Current Expression</param>
        // <param name="currentOptPos">Current Operator Location</param>
        // <param name="adjustOptPos">Adjusted Operator Position</param>
        // <returns> returns the adjusted operator</returns>
        public static string AdjustOperator(string currentOpt, string currentExp, ref int currentOptPos)
        {
            switch (currentOpt)
            {
                case "<":
                    if (currentExp.Substring(currentOptPos, 2) == "<=")
                    {
                        currentOptPos++;
                        return "<=";
                    }
                    if (currentExp.Substring(currentOptPos, 2) == "<>")
                    {
                        currentOptPos++;
                        return "<>";
                    }
                    return "<";

                case ">":
                    if (currentExp.Substring(currentOptPos, 2) == ">=")
                    {
                        currentOptPos++;
                        return ">=";
                    }
                    return ">";
                case "t":
                    if (currentExp.Substring(currentOptPos, 3) == "tan")
                    {
                        currentOptPos += 2;
                        return "tan";
                    }
                    return "error";
                case "a":
                    if (currentExp.Substring(currentOptPos, 4) == "atan")
                    {
                        currentOptPos += 3;
                        return "atan";
                    }
                    return "error";
                default:
                    return currentOpt;
            }
        }

        /// <summary>
        /// Converts an operator to a specified type
        /// </summary>
        // <param name="opt">Operator</param>
        // <param name="isBinaryOperator">Is it a binary operator</param>
        // <returns> returns the specified operator type</returns>
        public static OperatorType ConvertOperator(string opt, bool isBinaryOperator)
        {
            switch (opt)
            {
                case "!": return OperatorType.NOT;
                case "+": return isBinaryOperator ? OperatorType.ADD : OperatorType.PS;
                case "^": return isBinaryOperator ? OperatorType.PWR : OperatorType.ERR;
                case "-": return isBinaryOperator ? OperatorType.SUB : OperatorType.NS;
                case "*": return isBinaryOperator ? OperatorType.MUL : OperatorType.ERR;
                case "/": return isBinaryOperator ? OperatorType.DIV : OperatorType.ERR;
                case "%": return isBinaryOperator ? OperatorType.MOD : OperatorType.ERR;
                case "<": return isBinaryOperator ? OperatorType.LT : OperatorType.ERR;
                case ">": return isBinaryOperator ? OperatorType.GT : OperatorType.ERR;
                case "<=": return isBinaryOperator ? OperatorType.LE : OperatorType.ERR;
                case ">=": return isBinaryOperator ? OperatorType.GE : OperatorType.ERR;
                case "<>": return isBinaryOperator ? OperatorType.UT : OperatorType.ERR;
                case "=": return isBinaryOperator ? OperatorType.ET : OperatorType.ERR;
                case "&": return isBinaryOperator ? OperatorType.AND : OperatorType.ERR;
                case "|": return isBinaryOperator ? OperatorType.OR : OperatorType.ERR;
                case ",": return isBinaryOperator ? OperatorType.CA : OperatorType.ERR;
                case "@": return isBinaryOperator ? OperatorType.END : OperatorType.ERR;
                default: return OperatorType.ERR;
            }
        }

        /// <summary>
        /// Converts an operator to a specified type
        /// </summary>
        // <param name="opt">Operator</param>
        // <returns> returns the specified operator type</returns>
        public static OperatorType ConvertOperator(string opt)
        {
            switch (opt)
            {
                case "!": return OperatorType.NOT;
                case "+": return OperatorType.ADD;
                case "-": return OperatorType.SUB;
                case "*": return OperatorType.MUL;
                case "/": return OperatorType.DIV;
                case "%": return OperatorType.MOD;
                case "^": return OperatorType.PWR;
                case "<": return OperatorType.LT;
                case ">": return OperatorType.GT;
                case "<=": return OperatorType.LE;
                case ">=": return OperatorType.GE;
                case "<>": return OperatorType.UT;
                case "=": return OperatorType.ET;
                case "&": return OperatorType.AND;
                case "|": return OperatorType.OR;
                case ",": return OperatorType.CA;
                case "@": return OperatorType.END;
                case "tan": return OperatorType.TAN;
                case "atan": return OperatorType.ATAN;
                default: return OperatorType.ERR;
            }
        }

        /// <summary>
        /// If the operator is a binary operator, there is a problem with this method, don't use it yet.
        /// </summary>
        // <param name="tokens"> Syntax unit stack</param>
        // <param name="operators">Operator stack</param>
        // <param name="currentOpd">Current number of operations</param>
        /// <returns>yes to true, no to false</returns>
        public static bool IsBinaryOperator(ref Stack<object> tokens, ref Stack<Operator> operators, string currentOpd)
        {
            if (currentOpd != "")
            {
                return true;
            }
            else
            {
                object token = tokens.Peek();
                if (token is Operand)
                {
                    if (operators.Peek().Type != OperatorType.LB)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    if (((Operator)token).Type == OperatorType.RB)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }

        /// <summary>
        /// Operator priority comparison
        /// </summary>
        /// <param name="optA">Operator type A</param>
        /// <param name="optB">Operator type B</param>
        /// <returns>-1-Lower; 0-Equal; 1-Higher</returns>
        public static int ComparePriority(OperatorType optA, OperatorType optB)
        {
            if (optA == optB)
            {
                //A and B have equal priority
                return 0;
            }
            //*,/,%
            if ((optA >= OperatorType.MUL && optA <= OperatorType.MOD) &&
                (optB >= OperatorType.MUL && optB <= OperatorType.MOD))
            {
                return 0;
            }
            //+,-
            if ((optA >= OperatorType.ADD && optA <= OperatorType.SUB) &&
                (optB >= OperatorType.ADD && optB <= OperatorType.SUB))
            {
                return 0;
            }
            //<,<=,>,>=
            if ((optA >= OperatorType.LT && optA <= OperatorType.GE) &&
                (optB >= OperatorType.LT && optB <= OperatorType.GE))
            {
                return 0;
            }
            //=,<>
            if ((optA >= OperatorType.ET && optA <= OperatorType.UT) &&
                (optB >= OperatorType.ET && optB <= OperatorType.UT))
            {
                return 0;
            }
            //trigonometric function
            if ((optA >= OperatorType.TAN && optA <= OperatorType.ATAN) &&
                    (optB >= OperatorType.TAN && optB <= OperatorType.ATAN))
            {
                return 0;
            }

            if (optA < optB)
            {
                //Priority A over B
                return 1;
            }

            //Priority A is lower than B
            return -1;

        }
    }
}
