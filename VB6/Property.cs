using System;
using System.Collections.Generic;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
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

        public Module TargetModule { get; set; }

        public Property(Module targetModule, string comment) : base(targetModule.Parent, null)
        {
            Block = new Block(this);
            Parameters = new List<Parameter>();
            TargetModule = targetModule;
            Parent = TargetModule;
            Comment = comment;
        }

        public void ConvertParts(IAmAProperty sourceProperty)
        {
            throw new NotImplementedException("Cast upward");
        }

        public override string GenerateCode()
        {
            throw new NotImplementedException("Cast upward");
        }
    }
}