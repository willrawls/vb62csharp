namespace MetX.VB6ToCSharp
{
    public interface IAmAProperty : IGenerate
    {
        string Comment { get; set; }
        int Indent { get; set; }
        string Name { get; set; }
        string Scope { get; set; }
        string Type { get; set; }
        bool Valid { get; set; }
        string Value { get; set; }

        void ConvertSourcePropertyParts(IAmAProperty sourceProperty);
    }
}