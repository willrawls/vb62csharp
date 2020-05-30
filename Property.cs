using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class Property : IAmAProperty
    {
        public string Direction;
        public List<string> LineList;

        public string Comment { get; set; }
        public string Name { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public string Scope { get; set; }
        public int Indent { get; set; }

        public List<Parameter> Parameters { get; set; }

        public Property(int parentIndent)
        {
            LineList = new List<string>();
            Parameters = new List<Parameter>();
            Indent = parentIndent + 1;
        }

        public string GenerateTargetCode()
        {
            throw new System.NotImplementedException();
        }

        public void ConvertSourcePropertyParts(IAmAProperty sourceProperty)
        {
            throw new System.NotImplementedException();
        }
    }
}