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
            ResetIndent(Indent);

            var result = new StringBuilder();

            if (Line.IsNotEmpty())
                result.AppendLine(Indentation + Line);

            if (Before.IsNotEmpty() && Children.IsNotEmpty())
                result.AppendLine(Indentation + Before);

            if (Children.IsNotEmpty())
                foreach (var codeLine in Children)
                    result.Append(codeLine.GenerateCode());

            if (After.IsNotEmpty() && Children.IsNotEmpty())
                result.AppendLine(Indentation + After);

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
                    ((Block) codeLine).SetupChildren();
            }
        }

        public override string ToString()
        {
            return Line;
        }

        public override bool Equals(object blockObject)
        {
            if (!(blockObject is Block))
                return false;

            var otherBlock = (Block) blockObject;

            if (Line != otherBlock.Line)
                return false;
            if (Children.Count != otherBlock.Children.Count)
                return false;
            for (var index = 0; index < Children.Count; index++)
            {
                var child = Children[index];
                if (otherBlock.Children[index].Line != child.Line)
                    return false;
            }

            return true;
        }
    }
}