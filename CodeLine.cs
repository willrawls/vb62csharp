using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CodeLine : ICodeLine
    {
        public int Indent { get; set; }
        public string Line { get; set; }
        public ICodeLine Parent { get; set; }

        public CodeLine(ICodeLine parent, string line, int indent = 0)
        {
            Parent = parent;
            Line = line;
            Indent = indent;
        }

        public virtual string GenerateCode()
        {
            return Line ?? "";
        }

        public virtual bool IsEmpty() => Line.IsEmpty();

        public virtual bool IsNotEmpty() => Line.IsNotEmpty();
    }
}