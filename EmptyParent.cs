namespace MetX.VB6ToCSharp
{
    public class EmptyParent : Indentifier, ICodeLine
    {
        public EmptyParent(int indent = 0)
        {
            Indent = indent;
        }

        public string GenerateCode()
        {
            return string.Empty;
        }

        public int Indent { get; }
        public string Line { get; set; }
        public ICodeLine Parent { get; set; }

        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;
    }
}