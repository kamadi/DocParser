using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocParser
{
    class Util
    {
        public static string convert(string s)
        {
            return s.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
        }
    }
}
