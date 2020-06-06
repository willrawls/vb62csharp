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
            Indent = parent.Indent + 1;
            ParameterList = new List<Parameter>();
            LinesAfter = new Block(this);
        }

        public override string GenerateCode()
        {
            var indentation = Tools.Indent(Indent);
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty())
                return indentation + $"{finalPartType};";
            else
            {
                // Line = $"{finalPartType}";

                var result = new StringBuilder();
                result.AppendLine(indentation + $"{finalPartType}"); // Line);
                result.AppendLine(indentation + (Before ?? "{"));

                if(Line.IsNotEmpty())
                    result.AppendLine(indentation + Line);

                foreach (ICodeLine block in Blocks.Where(block => block.Line.IsNotEmpty()))
                {
                    var generatedCode = block.GenerateCode();
                    result.Append(generatedCode);
                }

                foreach (var blockAfter in LinesAfter.Blocks.Where(line => line.IsNotEmpty()))
                {
                    var generatedCode = blockAfter.GenerateCode();
                    result.Append(generatedCode);
                }

                result.AppendLine(indentation + (After ?? "}"));

                var code = result.ToString();
                return code;
            }
        }
    }
}