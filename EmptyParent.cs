namespace MetX.VB6ToCSharp
{
    public class EmptyParent : ICodeLine // : AbstractBlock
    {
        public EmptyParent(int indent = 1)
        {
            Indent = indent;
        }

        public string GenerateCode()
        {
            return "";
        }

        public int Indent { get; set; }
        public string Line { get; set; }
        public ICodeLine Parent { get; set; }
        public bool IsEmpty() => true;
        public bool IsNotEmpty() => false;
    }
}