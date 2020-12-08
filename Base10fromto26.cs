using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExel
{
    public class Base10fromto26
    {
        public Base10fromto26() { }

        public string From10To26 (int num)
        {

            string ans = "";
            while (num > 25)
            {
                ans = ans + ((char)('A' + num/26-1));
                num %= 26;
            }
            ans = ans + ((char)('A' + num)); ;
            return ans;
        }

        public int From26To10 (string num)
        {
            int ans = 0;
            char[] numArray = num.ToCharArray();
            int len = numArray.Length;
            int mult = 1;
            for (int i = len-1; i >= 0; i--)
            {
                ans += ((int)numArray[i] - (int)('A')+1) *mult;
                mult *= 26;
            }

            return ans-1;
        }
    }
}
