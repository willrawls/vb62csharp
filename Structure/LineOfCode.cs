using System;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.VB6;

// ReSharper disable InconsistentNaming

namespace MetX.VB6ToCSharp.Structure
{
    public class LineOfCode : Indentifier, ICodeLine
    {
        public string Before { get; set; }
        public string Line { get; set; }
        public string After { get; set; }
        public ICodeLine Parent { get; set; }

        public LineOfCode(ICodeLine parent, string line = null)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            Line = line;
        }

        public LineOfCode(string line)
        {
            Line = line;
        }

        public virtual string GenerateCode(int indent)
        {
            ResetIndent(indent);
            return $"{Indentation}{Line ?? ""}\r\n";
        }

        public virtual bool IsEmpty() => Line.IsEmpty();

        public virtual bool IsNotEmpty() => Line.IsNotEmpty();

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}