using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.Interface
{
    public interface IAmAProperty : IGenerate
    {
        string Comment { get; set; }
        int Indent { get; }
        string Name { get; set; }
        string Scope { get; set; }
        string Type { get; set; }
        bool Valid { get; set; }
        string Value { get; set; }

        Module TargetModule { get; set; }

        void ConvertParts(IAmAProperty sourceProperty);
        void ResetIndent(int indentLevel);
    }
}