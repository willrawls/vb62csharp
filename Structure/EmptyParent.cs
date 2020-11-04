using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class EmptyParent : Indentifier, ICodeLine
    {
        public EmptyParent(int indent = 0)
        {
            Indent = indent;
        }

        public string GenerateCode()
        {
            return string.Empty;
        }

        public int Indent { get; }
        public string Before { get; set; }
        public string Line { get; set; }
        public string After { get; set; }
        public ICodeLine Parent { get; set; }

        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;
    }
}