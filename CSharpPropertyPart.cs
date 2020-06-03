using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpPropertyPart : AbstractCodeBlock, IGenerate, IHaveCodeBlockParent
    {
        public AbstractCodeBlock BottomLineList;
        public bool Encountered;
        public AbstractCodeBlock LineList;
        public List<Parameter> ParameterList;
        public PropertyPartType PartType;

        public CSharpPropertyPart(IHaveCodeBlockParent parent, PropertyPartType propertyPartType)
        {
            Parent = parent;
            PartType = propertyPartType;
            Indent = 2;
            //LineList = new CodeBlock(this);
            //BottomLineList = new CodeBlock(this);
            ParameterList = new List<Parameter>();
        }

        public string GenerateCode()
        {
            var indentation = Tools.Indent(Parent.Indent + 1);
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty())
                return indentation + $"{finalPartType};";
            else
            {
                Line = $"{indentation}{finalPartType}";

                var result = new StringBuilder();
                result.AppendLine(indentation + Line);
                result.AppendLine(indentation + (Before ?? "{"));

                if (LineList?.Children != null)
                {
                    foreach (var child in LineList.Children)
                    {
                        if (child.Line.IsNotEmpty())
                        {
                            result.AppendLine(child.GenerateCode());
                        }
                    }
                }

                if (BottomLineList?.Children != null)
                {
                    foreach (var child in BottomLineList.Children)
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