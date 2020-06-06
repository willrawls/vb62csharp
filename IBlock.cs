using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public interface IBlock : ICodeLine
    {
        string After { get; set; }
        string Before { get; set; }
        List<ICodeLine> Blocks { get; set; }
    }
}