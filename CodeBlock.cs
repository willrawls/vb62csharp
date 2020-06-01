using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CodeBlock : IGenerate
    {
        public string After;
        public string Before;
        public List<CodeBlock> Children;
        public int Indent;
        public string Line;
        public CodeBlock Parent;

        public CodeBlock(CodeBlock parent, string line = null, string before = "{",
            List<CodeBlock> children = null, string after = "}", int indent = 1)
        {
            SetupBlock(parent, line, before, children, after, indent);
        }

        public CodeBlock(string line = null, string before = "{", List<CodeBlock> children = null,
            string after = "}", int indent = 1)
        {
            SetupBlock(null, line, before, children, after, indent);
        }

        public CodeBlock(string line, List<CodeBlock> children)
        {
            SetupBlock(null, line, "{", children, "}", 1);
        }

        public CodeBlock(string line = "")
        {
            if (line == null)
                throw new InvalidEnumArgumentException("line is required");
            SetupBlock(null, line, "{", null, "}", 1);
        }

        public string GenerateCode()
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
                    result.Append(child.GenerateCode());

                if (After.IsNotEmpty())
                    result.AppendLine(indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public void SetParent(CodeBlock parent)
        {
            Parent = parent;
        }

        public void SetupBlock(CodeBlock parent, string line, string before,
            List<CodeBlock> children,
            string after, int indent)
        {
            Indent = indent;
            Before = before;
            After = after;
            Line = line;
            Parent = parent;
            Children = children;
            SetupChildren(indent);
        }

        public void SetupChildren(int indent)
        {
            if (indent < 0)
                Indent = Parent?.Indent ?? 1;
            else
                Indent = indent;

            if (Children.IsEmpty()) return;

            foreach (var child in Children)
            {
                child.SetParent(this);
                child.SetupChildren(Indent + 1);
            }
        }

        public override string ToString()
        {
            return Line;
        }
    }
}