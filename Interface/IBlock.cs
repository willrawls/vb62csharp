using System.Collections.Generic;

namespace MetX.VB6ToCSharp.Interface
{
    public interface IBlock : ICodeLine
    {
        string After { get; set; }
        string Before { get; set; }
        List<ICodeLine> Children { get; set; }
    }
}