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

        public int Indent { get; set; }
        public string Before { get; set; }
        public string Line { get; set; }
        public string After { get; set; }
        public ICodeLine Parent { get; set; }

        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;

        public void ResetIndent()
        {
            if (Parent == null)
                return;

            Indent = Parent.Indent + 1;
        }
    }
}