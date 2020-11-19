namespace MetX.VB6ToCSharp.Interface
{
    public interface ICodeLine : IGenerate
    {
        string Before { get; set; }
        string Line { get; set; }
        string After { get; set; }


        ICodeLine Parent { get; set; }

        int Indent { get; set;  }
        string Indentation { get; }
        string SecondIndentation { get; }

        bool IsEmpty();
        bool IsNotEmpty();

        void ResetIndent(int indentLevel);
    }
}