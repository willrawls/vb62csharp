namespace MetX.VB6ToCSharp
{
    public interface ICodeLine : IGenerate
    {
        int Indent { get; set; }
        string Line { get; set; }
        ICodeLine Parent { get; set; }

        bool IsEmpty();

        bool IsNotEmpty();
    }
}