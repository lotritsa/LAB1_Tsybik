using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public class MyTable
    {
        public Hashtable values, expressions;
        public Hashtable Values
        {
            get
            {
                return values;
            }
        }
        public Hashtable Expressions
        {
            get
            {
                return expressions;
            }
        }
        static MyTable instance;
        private MyTable()
        {
            values = new Hashtable();
            expressions = new Hashtable();
        }
        public static MyTable GetInstance()
        {
            if (instance == null)
                instance = new MyTable();
            return instance;
        }
        public void AddExpression (string cell, string expression)
        {
            if (expressions.Contains(cell))
            {
                expressions[cell] = expression;
                return;
            }
            expressions.Add(cell, expression);
            values.Add(cell, "0");
        }
        public void AddValue(string cell, string value)
        {
            if (values.Contains(cell))
            {
                values[cell] = value;
                return;
            }
            values.Add(cell, value);
        }
        public void DeleteHash(string key)
        {
            expressions.Remove(key);
            values.Remove(key);
        }
        public void AddBoth(string cell, string expression, string value)
        {
            AddExpression(cell, expression);
            AddValue(cell, value);
        }

        public void ReCalculateAll(Parser converter)
        {
            restart:

            foreach (DictionaryEntry pair in values)
            {
                string preValue = values[pair.Key].ToString();
                string newValue = converter.Calculate(expressions[pair.Key].ToString());
                if (preValue == newValue)
                {
                    continue;
                }
                values[pair.Key] = newValue;
                goto restart;
            }
        }

        IDictionary<string, string> map = new Dictionary<string, string>();
        IDictionary<string, string> cl = new Dictionary<string, string>();
        IDictionary<string, string> p = new Dictionary<string, string>();
        bool dfs_prov (string v)
        {
            cl[v] = "1";
            for (int i = 0; i<map[v].Split('!').Count(); i++)
            {
                string to = map[v].Split('!')[i];
                if (cl[to] == "0")
                {
                    p[to] = v;
                    if (dfs_prov(to)) return true;
                }
                else if (cl[to] == "1")
                {
                    return true;
                }
            }
            cl[v] = "2";
            return false;
        }

        public void ReCalculate(string cell, Parser converter)
        {
            /*
            // проверка на цикл
            map.Clear();
            cl.Clear();
            p.Clear();
            foreach (DictionaryEntry pair in expressions)
            {
                string index = pair.Key.ToString();
                string s = pair.Value.ToString();
                string new_s = "";
                string rowNum = "";
                string columnNum = "";
                for (int j = 0; j < s.Length; j++)
                {
                    if (s[j] >= 'A' && s[j] <= 'Z')
                    {
                        while (j < s.Length && s[j] >= 'A' && s[j] <= 'Z')
                        {
                            columnNum = columnNum + s[j];
                            j++;
                        }
                        while (j < s.Length && s[j] >= '0' && s[j] <= '9')
                        {
                            rowNum = rowNum + s[j];
                            j++;
                        }
                        new_s = new_s + columnNum + rowNum + '!';
                        cl[columnNum+rowNum] = "0";
                        p[columnNum+rowNum] = "-1";
                    }

                }
                if (new_s.Length > 1) new_s = new_s.Substring(0, new_s.Length - 1);
                map[index] = new_s;
                cl[index] = "0";
                p[index] = "-1";

            }
            foreach (DictionaryEntry pair in expressions)
            {
                string index = pair.Key.ToString();
                if (dfs_prov(index))
                {
                    string message1 = "Collision!";
                    string caption1 = "Alert";
                    MessageBoxButtons buttons1 = MessageBoxButtons.OK;
                    MessageBox.Show(message1, caption1, buttons1);
                        return;
                }
            }

            */
            //
            values[cell] = converter.Calculate(expressions[cell].ToString());
            restart:

            foreach (DictionaryEntry pair in values)
            {
                string preValue = values[pair.Key].ToString();
                string newValue = converter.Calculate(expressions[pair.Key].ToString());
                if (preValue == newValue)
                {
                    continue;
                }
                values[pair.Key] = newValue;
                goto restart;
            }
        }

    }
}
