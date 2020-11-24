using MetX.Library;
using MetX.VB6ToCSharp.CSharp;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    /// TODO - Add class summary
    /// </summary>
    public class Variable : Indentifier
    {
        public string Comment;
        public string Name;

        public string Scope;
        public string Type;

        public string GenerateCode(int indentLevel)
        {
            var result = Tools.Indent(indentLevel)
                + "public " // Scope
                + (Type.IsEmpty() ? "unknown" : Type) + " "
                + Name + ";";
            return result;
        }

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}