using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CodeBlock : AbstractCodeBlock
    {
        public CodeBlock(AbstractCodeBlock parent, string line = null, List<AbstractCodeBlock> children = null, string before = "{", string after = "}")
        {
            SetupBlock(parent, line, before, children, after);
        }

        public override string GenerateCode()
        {
            var indentation = Tools.Indent(Indent);
            var result = new StringBuilder();

            if (Line.IsNotEmpty())
                result.AppendLine(indentation + Line);

            if (!Children.IsEmpty())
            {
                if (Before.IsNotEmpty())
                    result.AppendLine(indentation + Before);

                foreach (var child in Children)
                    result.Append(((CodeBlock)child).GenerateCode());

                if (After.IsNotEmpty())
                    result.AppendLine(indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public string GenerateCode(AbstractCodeBlock parent)
        {
            throw new NotImplementedException();
        }

        public void SetupBlock(AbstractCodeBlock parent, string line, string before, List<AbstractCodeBlock> children, string after)
        {
            Parent = parent;
            Line = line;
            Before = before;
            After = after;
            Children = children ?? new List<AbstractCodeBlock>();
            SetupChildren();
        }

        public void SetupChildren()
        {
            Indent = Parent?.Indent + 1 ?? 1;

            if (Children.IsEmpty()) return;

            foreach (var child in Children)
            {
                child.Parent = this;
                if (Parent is CodeBlock)
                    ((CodeBlock)child).SetupChildren();
            }
        }

        public override string ToString()
        {
            return Line;
        }
    }
}