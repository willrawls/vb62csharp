using System;
using System.Collections.Generic;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public abstract class AbstractCodeBlock : CodeLine, ICodeBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";
        public List<ICodeLine> Children { get; set; }

        // ReSharper disable once PublicConstructorInAbstractClass
        public AbstractCodeBlock(ICodeLine parent, string line, int indent = 0) : base(parent, line, indent)
        {
        }

        public override bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }
}