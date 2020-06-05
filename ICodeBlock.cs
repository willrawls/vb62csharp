using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public interface ICodeBlock : ICodeLine
    {
        string After { get; set; }
        string Before { get; set; }
        List<ICodeLine> Children { get; set; }
    }
}