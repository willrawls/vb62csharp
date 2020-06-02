using System;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class ControlProperty : IAmAProperty, IGenerate
    {
        public List<IAmAProperty> PropertyList;
        public string Comment { get; set; }

        public int Indent { get; set; }
        public string Name { get; set; }

        public string Scope { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }

        public string Value { get; set; }

        public ControlProperty(int parentIndent)
        {
            PropertyList = new List<IAmAProperty>();
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

        public string GenerateCode(CodeBlock parent)
        {
            throw new NotImplementedException();
        }
    }
}