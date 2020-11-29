using System;
using System.CodeDom;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public abstract class AbstractBlock : LineOfCode, IBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";

        public CodeLineList Children { get; set; }

        protected int AbstractBlockResetIndentsRecursively(int startingIndentLevel)
        {
            foreach(var child in Children)
                child.ResetIndent(startingIndentLevel + 1);
            return startingIndentLevel;
        }

        protected AbstractBlock(ICodeLine parent, string line = null) 
            : base(parent, line)
        {
            Children = new CodeLineList();
            ResetIndentsRecursively = AbstractBlockResetIndentsRecursively;
        }

        public override bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }
}