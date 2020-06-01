using System;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class Property : IAmAProperty, IGenerate
    {
        public string Direction;
        public List<string> LineList;

        public string Comment { get; set; }

        public int Indent { get; set; }

        public string Name { get; set; }

        public List<Parameter> Parameters { get; set; }

        public string Scope { get; set; }

        public string Type { get; set; }

        public bool Valid { get; set; }

        public string Value { get; set; }

        public Property(int parentIndent)
        {
            LineList = new List<string>();
            Parameters = new List<Parameter>();
            Indent = parentIndent + 1;
        }

        public void ConvertSourcePropertyParts(IAmAProperty sourceProperty)
        {
            throw new NotImplementedException();
        }

        public string GenerateCode()
        {
            throw new NotImplementedException();
        }
    }
}