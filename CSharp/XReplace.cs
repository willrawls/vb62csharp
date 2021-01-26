namespace MetX.VB6ToCSharp.CSharp
{
    public class XReplace     // Four way replacement
    {
        public string X;
        public string Y;
        public string Z;
        public string A;

        public XReplace(string x, string y = null, string z = null, string a = null)
        {
            X = x;
            Y = y;
            Z = z;
            A = a;
        }
    }
}