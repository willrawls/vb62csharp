using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public class FourWayReplace     // Four way replacement
    {
        public string X;
        public string Y;
        public string Z;
        public string A;

        public FourWayReplace(string x, string y = null, string z = null, string a = null)
        {
            X = x;
            Y = y;
            Z = z;
            A = a;
        }
    }
}