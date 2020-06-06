using System;
using System.Collections.Generic;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public abstract class AbstractBlock : LineOfCode, IBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";
        public List<ICodeLine> Blocks { get; set; }

        // ReSharper disable once PublicConstructorInAbstractClass
        public AbstractBlock(ICodeLine parent, string line) : base(parent, line)
        {
            Blocks = new List<ICodeLine>();
        }

        public override bool IsEmpty() => Line.IsEmpty() && Blocks?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Blocks?.Count > 0;
    }
}