namespace MetX.VB6ToCSharp
{
    public interface IHaveCodeBlockParent
    {
        int Indent { get; set; }
        string Line { get; set; }
        IHaveCodeBlockParent Parent { get; set; }
    }
}