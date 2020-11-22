using System.Collections.Generic;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class CodeLines : List<ICodeLine>
    {
    }

    public abstract class AbstractBlock : LineOfCode, IBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";

        public CodeLines Children { get; set; }

        // ReSharper disable once PublicConstructorInAbstractClass
        public AbstractBlock(ICodeLine parent, string line = null) : base(parent, line)
        {
            Children = new CodeLines();
        }

        public override bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }
}