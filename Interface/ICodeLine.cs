namespace MetX.VB6ToCSharp.Interface
{
    public interface ICodeLine : IGenerate
    {
        string Line { get; set; }
        ICodeLine Parent { get; set; }

        int Indent { get; }
        string Indentation { get; }
        string SecondIndentation { get; }

        bool IsEmpty();

        bool IsNotEmpty();
    }
}