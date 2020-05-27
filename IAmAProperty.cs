namespace MetX.VB6ToCSharp
{
    public interface IAmAProperty
    {
        string Comment { get; set; }
        string Name { get; set; }
        bool Valid { get; set; }
        string Value { get; set; }
        string Type { get; set; }
        string Scope { get; set; }
        int Indent { get; set; }

        string GenerateTargetCode();
        void ParsePropertyParts(IAmAProperty sourceProperty);
    }
}