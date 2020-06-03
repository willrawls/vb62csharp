using System;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public class AbstractCodeBlock : IGenerate, IHaveCodeBlockParent
    {
        public string After = "}";
        public string Before = "{";
        public List<AbstractCodeBlock> Children;
        public int Indent { get; set; }
        public string Line { get; set; }
        public IHaveCodeBlockParent Parent { get; set; }

        public virtual string GenerateCode()
        {
            throw new NotImplementedException();
        }
    }
}