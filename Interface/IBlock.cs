using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.Interface
{
    public interface IBlock : ICodeLine
    {
        CodeLineList Children { get; set; }
    }
}