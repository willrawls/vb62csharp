using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpPropertyPart : AbstractBlock, IGenerate, ICodeLine
    {
        public AbstractBlock LinesAfter;

        //public AbstractCodeBlock BlockAtTop;
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
            else
            {
                var result = new StringBuilder();
                result.AppendLine(Indentation + $"{finalPartType}");
                if(Before.IsNotEmpty())
                   result.AppendLine(Indentation + Before);

                if(Line.IsNotEmpty())
                {
                    result.AppendLine(SecondIndentation + Line);
                    result.AppendLine(SecondIndentation + Before);
                }

                foreach (var codeLine in Children.Where(block => block.Line.IsNotEmpty()))
                {
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
}