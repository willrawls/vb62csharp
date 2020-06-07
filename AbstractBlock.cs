using System;
using System.Collections.Generic;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public abstract class AbstractBlock : LineOfCode, IBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";
        public List<ICodeLine> Children { get; set; }

        // ReSharper disable once PublicConstructorInAbstractClass
        public AbstractBlock(ICodeLine parent, string line) : base(parent, line)
        {
            Children = new List<ICodeLine>();
        }

        public override bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }
}