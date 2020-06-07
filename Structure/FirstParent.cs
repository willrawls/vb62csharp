using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class FirstParent : Indentifier, ICodeLine
    {
        public FirstParent(int indent = 0)
        {
            Indent = indent;
        }

        public string GenerateCode()
        {
            return string.Empty;
        }

        public int Indent { get; }
        public string Line { get; set; }
        public ICodeLine Parent { get; set; }

        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;
    }
}