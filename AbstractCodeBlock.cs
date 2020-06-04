using System;
using System.Collections.Generic;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public interface ICodeBlock : IGenerate, ICodeLine
    {
        bool IsEmpty();

        bool IsNotEmpty();
    }

    public interface ICodeLine
    {
        int Indent { get; set; }
        string Line { get; set; }
        IHaveCodeBlockParent Parent { get; set; }
    }

    public class AbstractCodeBlock : IGenerate, IHaveCodeBlockParent, ICodeBlock
    {
        public string After = "}";
        public string Before = "{";
        public List<AbstractCodeBlock> Children;
        public int Indent { get; set; }
        public string Line { get; set; }

        public IHaveCodeBlockParent Parent { get; set; }

        public virtual string GenerateCode()
        {
            throw new NotImplementedException();
        }

        public bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }

    public class CodeLine : ICodeLine
    {
        public int Indent { get; set; }
        public string Line { get; set; }
        public IHaveCodeBlockParent Parent { get; set; }

        public CodeLine(IHaveCodeBlockParent parent, string line, int indent = 0)
        {
            Parent = parent;
            Line = line;
            Indent = indent;
        }
    }
}