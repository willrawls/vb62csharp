using System.Collections.Generic;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class FirstParent : Indentifier, ICodeLine
    {
        public FirstParent(int indent = 0)
        {
            ResetIndent(indent);
        }

        public int Indent { get; set; }
        public string Before { get; set; }
        public string Line { get; set; }
        public string After { get; set; }
        public ICodeLine Parent { get; set; }
        public List<ICodeLine> Children { get; set; } = new List<ICodeLine>();

        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;

        public string GenerateCode(int indentLevel)
        {
            return string.Empty;
        }

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}