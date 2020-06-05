using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class ControlProperty : AbstractCodeBlock, IAmAProperty
    {
        public List<IAmAProperty> PropertyList;
        public string Comment { get; set; }
        public string Name { get; set; }
        public string Scope { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }

        public ControlProperty(ICodeLine parent, int parentIndent) : base(parent, "", 0)
        {
            PropertyList = new List<IAmAProperty>();
            Indent = parentIndent + 1;
        }

        public void ConvertSourcePropertyParts(IAmAProperty sourceProperty)
        {
            throw new NotImplementedException();
        }

        public override string GenerateCode()
        {
            throw new NotImplementedException();
        }
    }
}