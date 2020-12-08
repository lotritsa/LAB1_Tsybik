using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public class Parser
    {
        private List<string> acceptableOperators = new List<string>(new string[] {"(",")","+","-","/", "*", "mod", "div", "=", "<=", ">=", "<>", "<", ">", "not"});
        List<string> res;
        public Parser() { }

        private string getNextTerm(ref string str)
        {
            foreach (string op in acceptableOperators)
            {
                if (str.StartsWith(op))
                {
                    str = str.Substring(op.Length, str.Length - op.Length);
                    return op;
                }
            }

            int last = 0;
            while (last < str.Length && str[last] >= 'A' && str[last] <= 'Z')
            {
                last++;
            }
            while (last < str.Length && str[last] >= '0' && str[last] <= '9')
            {
                last++;
            }

            if (last == 0)
            {
                throw new Exception("Unrecognized character");
            }

            string term = str.Substring(0, last);
            str = str.Substring(term.Length, str.Length - term.Length);

            return term;
        }

        public void StringToExpression(string str)
        {

            res = new List<string>();
            Stack<string> strOperators = new Stack<string>();
            bool isUnary = true;

            while (!(str.Length == 0) && str[0] == ' ')
            {
                str = str.Substring(1, str.Length-1);
            }
            while (!(str.Length == 0))
            {
                string term = getNextTerm(ref str);
                if (isUnary && term.Equals("-"))
                {
                    while (!(str.Length == 0) && str[0] == ' ')
                    {
                        str = str.Substring(1, str.Length-1);
                    }
                    if (str.Length == 0)
                    {
                        throw new Exception("Incorrect expression: minus at the end");
                    }

                    string nextTerm = getNextTerm(ref str);
                    if (nextTerm[0] >= '0' && nextTerm[0] <= '9')
                    {
                        term = term + nextTerm;
                    } else if(nextTerm[0] >= 'A' && nextTerm[0] <= 'Z')
                    {
                        Cell CellVar = new Cell();
                        term = (-Convert.ToDouble(CellVar.GetValue(nextTerm))).ToString();
                    } else
                    {
                        throw new Exception("Unexpected token");
                    }
                }
                if (acceptableOperators.Contains(term))
                {
                    if (strOperators.Count != 0 && !term.Equals("("))
                    {
                        if (term.Equals(")"))
                        {
                            string s = strOperators.Pop();
                            while (s != "(")
                            {
                                res.Add(s);
                                s = strOperators.Pop();
                            }
                        }
                        else if (GetPriority(term) > GetPriority(strOperators.Peek()))
                        {
                            strOperators.Push(term);
                        }
                        else if (strOperators.Count > 0 && GetPriority(term) <= GetPriority(strOperators.Peek()))
                        {
                            res.Add(strOperators.Pop());
                            strOperators.Push(term);
                        }
                    }
                    else strOperators.Push(term);
                }
                else
                {
                    if (term[0] >= 'A' && term[0] <= 'Z')
                    {
                        Cell CellVar = new Cell();
                        term = CellVar.GetValue(term);
                    }
                    res.Add(term);
                }
                if (term.Equals("("))
                {
                    isUnary = true;
                } else
                {
                    isUnary = false;
                }

                while (!(str.Length == 0) && str[0] == ' ')
                {
                    str = str.Substring(1, str.Length-1);
                }
            }
            while (strOperators.Count != 0)
            {
                res.Add(strOperators.Pop());
            }
            res.ToArray();
        }
        public int GetPriority(string num)
        {
            switch (num)
            {
                case "(":
                case ")":
                    return 0;
                case ">":
                case "<":
                case "=":
                case ">=":
                case "<=":
                case "<>":
                case "+":
                case "-":
                    return 2;
                case "*":
                case "/":
                case "mod":
                case "div":
                case "not":
                    return 3;
                default:
                    return 4;

            }
        }

        public string Calculate(string cur)
        {
            
            //SpaceCreated(ref cur);
            StringToExpression(cur);
            Stack<string> stk = new Stack<string>();
            Queue<string> que = new Queue<string>(res);
            string str = que.Dequeue();
            while (que.Count >= 0)
            {
                if (!acceptableOperators.Contains(str))
                {
                    stk.Push(str);
                    if (que.Count == 0)
                    {
                        break;
                    }
                    str = que.Dequeue();
                }
                else
                {
                    double ans = 0;
                    try
                    {
                        switch (str)
                        {
                            case "+":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = a + b;
                                    break;
                                }
                            case "-":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = b - a;
                                    break;
                                }
                            case "*":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = a*b;
                                    break;
                                }
                            case "/":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = b / a;
                                    break;
                                }
                            case "mod":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = b % a;
                                    break;
                                }
                            case "div":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    ans = (int)(b / a);
                                    break;
                                }
                            case "=":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (a == b) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case ">":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (b > a) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case "<":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (b < a) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case ">=":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (b >= a) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case "<=":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (b <= a) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case "<>":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    double b = Convert.ToDouble(stk.Pop());
                                    if (a != b) ans = 1;
                                    else ans = 0;
                                    break;
                                }
                            case "not":
                                {
                                    double a = Convert.ToDouble(stk.Pop());
                                    if (a == 0) ans = 1;
                                    else ans = 0;
                                    break;
                                }

                        }
                    }
                    catch
                    {

                    }
                    stk.Push(ans.ToString());
                    if (que.Count > 0)
                        str = que.Dequeue();
                    else
                        break;
                }
            }

            return stk.Pop();
        }

    }
}

