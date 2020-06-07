using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class Block : AbstractBlock
    {
        public Block(ICodeLine parent, string line = null)
            : base(parent, line)
        {
           SetupBlock(parent, line, "{", "}");
        }

        public override string GenerateCode()
        {
            var result = new StringBuilder();

            if (Line.IsNotEmpty())
                result.AppendLine(Indentation + Line);

            if (Children.IsNotEmpty())
            {
                if (Before.IsNotEmpty())
                    result.AppendLine(Indentation + Before);

                foreach (var codeLine in Children)
                    result.Append(codeLine.GenerateCode());

                if (After.IsNotEmpty())
                    result.AppendLine(Indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public void SetupBlock(ICodeLine parent, string line, string before, string after)
        {
            Parent = parent;
            Line = line;
            Before = before;
            After = after;
            SetupChildren();
        }

        public void SetupChildren()
        {
            if (Children.IsEmpty()) return;

            foreach (var codeLine in Children)
            {
                codeLine.Parent = this;
                if (Parent is AbstractBlock)
                    ((Block)codeLine).SetupChildren();
            }
        }

        public override string ToString()
        {
            return Line;
        }
    }
}