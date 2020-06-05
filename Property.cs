using System;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class Property : AbstractBlock, IAmAProperty
    {
        public Block Block;
        public string Direction;
        public string Comment { get; set; }
        public string Name { get; set; }
        public List<Parameter> Parameters { get; set; }
        public string Scope { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }

        public Property(ICodeLine parent) : base(parent, null)
        {
            Block = new Block(this);
            Parameters = new List<Parameter>();
            Indent = parent.Indent + 1;
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