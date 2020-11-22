using System.Collections.Generic;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.Interface
{
    public interface IBlock : ICodeLine
    {
        CodeLines Children { get; set; }
    }
}