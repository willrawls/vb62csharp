using System;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

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

        public string Final => GenerateCode();

        public virtual string GenerateCode()
        {
            return $"{Indentation}{Line ?? ""}\r\n";
        }

        public virtual bool IsEmpty() => Line.IsEmpty();

        public virtual bool IsNotEmpty() => Line.IsNotEmpty();

        protected Func<int, int> ResetIndentsRecursively;

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = ResetIndentsRecursively?.Invoke(indentLevel) ?? indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}