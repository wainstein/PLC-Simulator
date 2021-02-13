using System;
using System.Collections.Generic;
using PLCTools.Models;

namespace PLCTools.Service.MathPaser
{
    /// Reverse Polish Notation					
    public class RPN
    {
        Stack<object> m_tokens = new Stack<object>();
        public string Expression { get; set; } = "";
        /// Finalize reverse polish notation stack		
        public Stack<object> Tokens
        {
            get { return m_tokens; }
            set { m_tokens = value; }
        }
        readonly List<string> m_Operators = new List<string>(new string[]{
         "(","tan",")","atan","!","*","/","%","+","-","<",">","<>",">=","<=","=","&","|",",","@","^"});   //Operator allowed
        private bool IsMatching(string exp)
        {
            string opt = "";    //temp storage " ' # (	
            for (int i = 0; i < exp.Length; i++)
            {
                string chr = exp.Substring(i, 1);   //read each byte				
                if ("\"'#".Contains(chr))   //The current character is a type of double quotation mark, single quotation mark, or well mark
                {
                    if (opt.Contains(chr))  //This character has been read before.	
                    {
                        opt = opt.Remove(opt.IndexOf(chr), 1);  //Removes the character that was previously read as a match
                    }
                    else
                    {
                        opt += chr;  //The first time the character is read, save
                    }
                }
                else if ("()".Contains(chr))    //left and right brackets
                {
                    if (chr == "(")
                    {
                        opt += chr;
                    }
                    else if (chr == ")")
                    {
                        if (opt.Contains("("))
                        {
                            opt = opt.Remove(opt.IndexOf("("), 1);
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            return (opt == "");
        }
        /// Finding an operator position from an expression			
        /// <param name="exp">Expression</param>					
        /// <param name="findOpt">Operators to look for</param>					
        /// <returns>Returns the operator position</returns>					
        private int FindOperator(string exp, string findOpt)
        {
            string opt = "";
            for (int i = 0; i < exp.Length; i++)
            {
                string chr = exp.Substring(i, 1);
                if ("\"'#".Contains(chr))//Ignore operators in double quotes, single quotes, and well symbols
                {
                    if (opt.Contains(chr))
                    {
                        opt = opt.Remove(opt.IndexOf(chr), 1);
                    }
                    else
                    {
                        opt += chr;
                    }
                }
                if (opt == "")
                {
                    if (findOpt != "")
                    {
                        if (findOpt == chr)
                        {
                            return i;
                        }
                    }
                    else
                    {
                        if (m_Operators.Exists(x => x.Contains(chr)))
                        {
                            return i;
                        }
                    }
                }
            }
            return -1;
        }
        public bool Parse(string exp)
        {
            Logging log = new Logging(exp);
            Expression = exp;
            try
            {
                m_tokens.Clear();//Clear math operation units stack
                if (exp.Trim() == "")//the expressions should have something												
                {
                    return false;
                }
                else if (!this.IsMatching(exp))//match the operands
                {
                    return false;
                }

                Stack<object> operands = new Stack<object>();       //Stack of Operands									
                Stack<Operator> operators = new Stack<Operator>();  //Stack of Operators											
                OperatorType optType = OperatorType.ERR;            //Type of Operator								
                string curOpd;                                      //Current Operand
                string curOpt = "";                                 //Current Operator
                int curPos = 0;                                     //Current Position			
                                                                    //int funcCount = 0;                                //Function Count		

                curPos = FindOperator(exp, "");

                exp += "@"; //End of the Operation											
                while (true)
                {
                    curPos = FindOperator(exp, "");

                    curOpd = exp.Substring(0, curPos).Trim();
                    curOpt = exp.Substring(curPos, 1);



                    //Stores the current operands to the operand stack
                    if (curOpd != "")
                    {
                        operands.Push(new Operand(curOpd, curOpd));
                    }

                    //If the current operator is the end operator, the loop is stopped
                    if (curOpt == "@")
                    {
                        break;
                    }
                    //If the current operator is in left brackets, it is stored directly on the stack.
                    if (curOpt == "(")
                    {
                        operators.Push(new Operator(OperatorType.LB, "("));
                        exp = exp.Substring(curPos + 1).Trim();
                        continue;
                    }

                    //If the current operator is in right brackets, pop up the operators in 
                    //the operator stack and store them in the operator stack until you meet 
                    //the left brackets, then discard the left brackets.
                    if (curOpt == ")")
                    {
                        while (operators.Count > 0)
                        {
                            if (operators.Peek().Type != OperatorType.LB)
                            {
                                operands.Push(operators.Pop());
                            }
                            else
                            {
                                operators.Pop();
                                break;
                            }
                        }
                        exp = exp.Substring(curPos + 1).Trim();
                        continue;
                    }
                    //Adjustment Operators
                    curOpt = Operator.AdjustOperator(curOpt, exp, ref curPos);
                    optType = Operator.ConvertOperator(curOpt);
                    //If the operator stack is empty, or if the top of the operator stack is in 
                    //left brackets, the current operator is stored directly into the operator stack.
                    if (operators.Count == 0 || operators.Peek().Type == OperatorType.LB)
                    {
                        operators.Push(new Operator(optType, curOpt));
                        exp = exp.Substring(curPos + 1).Trim();
                        continue;
                    }
                    //If the current operator has a higher priority than the operator at the 
                    //top of the stack, the current operator will be stored directly into the stack.
                    if (Operator.ComparePriority(optType, operators.Peek().Type) > 0)
                    {
                        operators.Push(new Operator(optType, curOpt));
                    }
                    else
                    {
                        //If the current operator has a lower or equal priority than 
                        //the operator at the top of the operator stack, output the 
                        //top-of-stack operator to the operand stack until the top-of-stack 
                        //operator is lower than (not including equal to) that operator's priority.
                        //or operator stack top-of-stack operator in left brackets.
                        //and presses the current operator into the operator stack.
                        while (operators.Count > 0)
                        {
                            if (Operator.ComparePriority(optType, operators.Peek().Type) <= 0 && operators.Peek().Type != OperatorType.LB)
                            {
                                operands.Push(operators.Pop());

                                if (operators.Count == 0)
                                {
                                    operators.Push(new Operator(optType, curOpt));
                                    break;
                                }
                            }
                            else
                            {
                                operators.Push(new Operator(optType, curOpt));
                                break;
                            }
                        }

                    }
                    exp = exp.Substring(curPos + 1).Trim();
                }
                //Conversion complete, if there are still operators in the operator stack.
                //Then remove the operators from the operator stack in order until the operator stack is empty.
                while (operators.Count > 0)
                {
                    operands.Push(operators.Pop());
                }
                //Reorder the objects in the operands stack and output them to the final stack.												
                while (operands.Count > 0)
                {
                    m_tokens.Push(operands.Pop());
                }
                //log.Success();
                return true;
            }
            catch(Exception ex)
            {
                log.Fatal(ex);
                return false;
            }
        }
        public object Evaluate()
        {
            /*									
			  Reverse Polish Notation evaluation algorithm.	
			  1. Circulate scan grammar unit project.									
			  2. If the item scanned is an operand, it is pressed into the operand stack and the next item is scanned。									
			  3. If the item scanned is a binary operator, the operation is performed on the top two operands of the stack.									
			  4. If the item scanned is a one-dimensional operator, the operation is performed on the top operand of the stack.									
			  5. Presses the result of the calculation back into the stack.									
			  6. Repeat steps 2-5 and the resultant value is in the stack.									
			*/
            Logging log = new Logging(Expression);
            try
            {
                if (m_tokens.Count == 0) return null;

                object value = null;
                Stack<Operand> opds = new Stack<Operand>();
                Stack<object> pars = new Stack<object>();
                Operand opdA, opdB;

                foreach (object item in m_tokens)
                {
                    if (item is Operand operand)
                    {
                        //Parsing equations, replacing parameters
                        //If it is an operand, press into the operand stack.
                        opds.Push(operand);
                    }
                    else
                    {
                        switch (((Operator)item).Type)
                        {
                            #region Pwr,^,Power						
                            case OperatorType.PWR:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, Math.Pow(double.Parse(opdB.Value.ToString()), double.Parse(opdA.Value.ToString()))));
                                }
                                else
                                {
                                    throw new Exception("Powering must be number between two figure");
                                }
                                break;
                            #endregion
                            #region *,multiplication						
                            case OperatorType.MUL:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, double.Parse(opdB.Value.ToString()) * double.Parse(opdA.Value.ToString())));
                                }
                                else
                                {
                                    throw new Exception("Multiplication must be number between two figure");
                                }
                                break;
                            #endregion
                            #region /,division						
                            case OperatorType.DIV:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, double.Parse(opdB.Value.ToString()) / double.Parse(opdA.Value.ToString())));
                                }
                                else
                                {
                                    throw new Exception("Division must be number between two figure");
                                }
                                break;
                            #endregion
                            #region %,modulus						
                            case OperatorType.MOD:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, double.Parse(opdB.Value.ToString()) % double.Parse(opdA.Value.ToString())));
                                }
                                else
                                {
                                    throw new Exception("Modulus must be number between two figure");
                                }
                                break;
                            #endregion
                            #region +,Addition						
                            case OperatorType.ADD:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, double.Parse(opdB.Value.ToString()) + double.Parse(opdA.Value.ToString())));
                                }
                                else
                                {
                                    throw new Exception("Addition must be number between two figure");
                                }
                                break;
                            #endregion
                            #region -,subtraction						
                            case OperatorType.SUB:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, double.Parse(opdB.Value.ToString()) - double.Parse(opdA.Value.ToString())));
                                }
                                else
                                {
                                    throw new Exception("Substraction must be number between two figure");
                                }
                                break;
                            #endregion
                            #region tan,subtraction						
                            case OperatorType.TAN:
                                opdA = opds.Pop();
                                if (Operand.IsNumber(opdA.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, Math.Tan(double.Parse(opdA.Value.ToString()) * Math.PI / 180)));
                                }
                                else
                                {
                                    throw new Exception("Tangent must be number between two figure");
                                }
                                break;
                            #endregion
                            #region atan,subtraction						
                            case OperatorType.ATAN:
                                opdA = opds.Pop();
                                if (Operand.IsNumber(opdA.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, Math.Atan(double.Parse(opdA.Value.ToString()))));
                                }
                                else
                                {
                                    throw new Exception("Arctangent must be number between two figure");
                                }
                                break;
                            #endregion
                            #region Noteql,noeq,notequal						
                            case OperatorType.UT:
                                opdA = opds.Pop();
                                opdB = opds.Pop();
                                if (Operand.IsNumber(opdA.Value) && Operand.IsNumber(opdB.Value))
                                {
                                    opds.Push(new Operand(OperandType.NUMBER, (double.Parse(opdB.Value.ToString()) != double.Parse(opdA.Value.ToString())) ? 1 : 0));
                                }
                                else
                                {
                                    throw new Exception("Comparation must be number between two figure");
                                }
                                break;
                                #endregion
                        }
                    }
                }

                if (opds.Count == 1)
                {
                    value = opds.Pop().Value;
                }
                log.Success("Successful, Result: " + value);
                return value;
            }catch(Exception ex)
            {
                log.Fatal(ex);
                return null;
            }
        }
    }
}