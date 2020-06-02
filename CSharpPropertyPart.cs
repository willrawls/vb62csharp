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

        public bool IsEmpty => Line.Trim().Length == 0
                                               && LineList.Children.All(x => x.Line.Trim().IsEmpty())
                               && BottomLineList.Children.All(x => x.Line.Trim().IsEmpty());

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
            var firstIndent = Tools.Indent(Parent.Indent + 1);
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty)
                return firstIndent + $"{finalPartType};";

            Line = $"{firstIndent}{finalPartType}";
            var indentation = Tools.Indent(Indent);
            var childIndentation = Tools.Indent(Indent + 1);

            var result = new StringBuilder();
            result.AppendLine(indentation + Line);
            result.AppendLine(indentation + Before);

            foreach (var child in LineList.Children)
            {
                if (child.Line.IsNotEmpty() && child.Line.Length > 0)
                {
                    //string line = Massage.Now(child.Line);
                    result.AppendLine(child.GenerateCode());
                    //result.AppendLine(childIndentation + Line);
                }
            }

            var code = result.ToString();
            return code;
        }
    }
}