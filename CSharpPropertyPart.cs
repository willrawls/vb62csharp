using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpPropertyPart : AbstractCodeBlock, IGenerate, ICodeLine
    {
        public AbstractCodeBlock BlockAtBottom;

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
        }

        public string GenerateCode()
        {
            var indentation = Tools.Indent(Indent);
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty())
                return indentation + $"{finalPartType};";
            else
            {
                Line = $"{finalPartType}";

                var result = new StringBuilder();
                result.AppendLine(indentation + Line);
                result.AppendLine(indentation + (Before ?? "{"));

                if (Children != null)
                {
                    foreach (var child in Children)
                    {
                        if (child.Line.IsNotEmpty())
                        {
                            result.AppendLine(child.GenerateCode());
                        }
                    }
                }

                if (BlockAtBottom?.Children != null)
                {
                    foreach (var child in BlockAtBottom.Children)
                    {
                        if (child.IsNotEmpty())
                        {
                            result.AppendLine(child.GenerateCode());
                        }
                    }
                }
                result.AppendLine(indentation + (After ?? "}"));

                var code = result.ToString();
                return code;
            }
        }
    }
}