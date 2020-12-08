using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public class Cell
    {
        public Cell()
        {

        }
        public string GetValue(string curCell)
        {
            MyTable cur = MyTable.GetInstance();

            var kek = cur.Values[curCell];
            string test;
            if (kek == null)
            {
                test = "0";
            } else
            {
                test = cur.Values[curCell].ToString();
            }
            return test;
        }
        //
       
        //
    }
}
