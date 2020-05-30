using System;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class ControlProperty : IAmAProperty
    {
        public string Comment { get; set; }

        public string Name { get; set; }

        public bool Valid { get; set; }

        public string Value { get; set; }

        public string Type { get; set; }

        public string Scope { get; set; }

        public int Indent { get; set; }

        public List<IAmAProperty> PropertyList;

        public ControlProperty(int parentIndent)
        {
            PropertyList = new List<IAmAProperty>();
            Indent = parentIndent + 1;
        }

        public string GenerateTargetCode()
        {
            throw new NotImplementedException();
        }

        public void ConvertSourcePropertyParts(IAmAProperty sourceProperty)
        {
            throw new NotImplementedException();
        }
    }
}