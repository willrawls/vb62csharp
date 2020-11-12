using MetX.Library;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    /// TODO - Add class summary
    /// </summary>
    public class Variable
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
    }
}