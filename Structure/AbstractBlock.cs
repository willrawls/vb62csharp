using System;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public abstract class AbstractBlock : LineOfCode, IBlock
    {
        public string After { get; set; } = "}";
        public string Before { get; set; } = "{";

        public CodeLineList Children { get; set; }

        protected AbstractBlock(ICodeLine parent, Func<int, int> resetIndentsRecursively, string line = null) : base(parent, resetIndentsRecursively, line)
        {
            Parent = parent;
            ResetIndentsRecursively = resetIndentsRecursively;
            Children = new CodeLineList();
        }

        protected AbstractBlock(ICodeLine parent, string line = null) : base(parent, null, line)
        {
            Children = new CodeLineList();
        }

        public override bool IsEmpty() => Line.IsEmpty() && Children?.Count == 0;

        public override bool IsNotEmpty() => Line.IsNotEmpty() || Children?.Count > 0;
    }
}