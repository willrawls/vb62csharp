﻿using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class LineOfCode : ICodeLine
    {
        public int Indent { get; set; }
        public string Line { get; set; }
        public ICodeLine Parent { get; set; }

        public LineOfCode(ICodeLine parent, string line)
        {
            Parent = parent;
            Line = line;
            Indent = parent.Indent + 1;
        }

        public virtual string GenerateCode()
        {
            var indentation = Tools.Indent(Indent);
            return indentation + (Line ?? "") + "\r\n";
        }

        public virtual bool IsEmpty() => Line.IsEmpty();

        public virtual bool IsNotEmpty() => Line.IsNotEmpty();
    }
}