using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class Block : AbstractBlock
    {
        public static Block New(AbstractBlock parent, string line = "")
        {
            var block = new Block(parent, line);
            return block;
        }

        // public Block(AbstractBlock parent, string line = null,
        public Block(ICodeLine parent, string line = null,
            List<ICodeLine> lines = null,
            string before = "{",
            string after = "}",
            int indent = 1)
            : base(parent, line)
        {
            SetupBlock(parent, line, before, lines, after);
        }

        public override string GenerateCode()
        {
            var indentation = Tools.Indent(Indent);
            var result = new StringBuilder();

            if (Line.IsNotEmpty())
                result.AppendLine(indentation + Line);

            if (!Blocks.IsEmpty())
            {
                if (Before.IsNotEmpty())
                    result.AppendLine(indentation + Before);

                foreach (var line in Blocks)
                    result.Append(((Block)line).GenerateCode());

                if (After.IsNotEmpty())
                    result.AppendLine(indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public string GenerateCode(AbstractBlock parent)
        {
            throw new NotImplementedException();
        }

        public void SetupBlock(ICodeLine parent, string line, string before, List<ICodeLine> lines, string after)
        {
            Parent = parent;
            Line = line;
            Before = before;
            After = after;
            if (lines.IsNotEmpty()) 
                Blocks.AddRange(lines);
            SetupChildren();
        }

        public void SetupChildren()
        {
            Indent = Parent?.Indent + 1 ?? 1;

            if (Blocks.IsEmpty()) return;

            foreach (var line in Blocks)
            {
                line.Parent = this;
                if (Parent is Block)
                    ((Block)line).SetupChildren();
            }
        }

        public override string ToString()
        {
            return Line;
        }
    }
}