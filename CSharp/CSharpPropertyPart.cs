using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public class CSharpPropertyPart : AbstractBlock, IGenerate, ICodeLine
    {
        public AbstractBlock LinesAfter;

        public bool Encountered;

        public List<Parameter> ParameterList;
        public PropertyPartType PartType;

        public CSharpPropertyPart(ICodeLine parent, PropertyPartType propertyPartType) : base(parent, null)
        {
            Parent = parent;
            PartType = propertyPartType;
            ParameterList = new List<Parameter>();
            LinesAfter = new Block(this);
        }

        public override string GenerateCode()
        {
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty())
                return Indentation + $"{finalPartType};";

            var result = new StringBuilder();
            result.AppendLine(Indentation + $"{finalPartType}");
            if(Before.IsNotEmpty())
                result.AppendLine(Indentation + Before);

            if(Line.IsNotEmpty())
            {
                result.AppendLine(SecondIndentation + Line);
                // result.AppendLine(SecondIndentation + Before);
            }

            foreach (ICodeLine codeLine in Children.Where(block => block.Line.IsNotEmpty()))
            {
                codeLine.Before = codeLine.After = null;
                codeLine.ResetIndent(Indent + 1);
                var generatedCode = codeLine.GenerateCode();
                result.Append(
                    Line.IsEmpty() 
                        ? generatedCode 
                        : generatedCode.AddIndent());
            }

            foreach (var blockAfter in LinesAfter.Children.Where(line => line.IsNotEmpty()))
            {
                var generatedCode = blockAfter.GenerateCode();
                result.Append(
                    Line.IsEmpty() 
                        ? generatedCode 
                        : generatedCode.AddIndent());
            }

            if(Line.IsNotEmpty())
            {
                result.AppendLine(SecondIndentation + After);
            }

            if(After.IsNotEmpty())
                result.AppendLine(Indentation + After);

            var code = result.ToString();
            return code;
        }
    }
}