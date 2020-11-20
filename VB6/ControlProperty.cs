using System;
using System.Collections.Generic;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class ControlProperty : AbstractBlock, IAmAProperty
    {
        public List<IAmAProperty> PropertyList;
        public string Comment { get; set; }
        public string Name { get; set; }
        public string Scope { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }
        public Module TargetModule { get; set; }

        public ControlProperty(ICodeLine parent, int parentIndent) : base(parent, "")
        {
            PropertyList = new List<IAmAProperty>();
        }

        public void ConvertParts(IAmAProperty sourceProperty)
        {
            throw new NotImplementedException();
        }

        public override string GenerateCode(int indentLevel)
        {
            throw new NotImplementedException();
        }
    }
}