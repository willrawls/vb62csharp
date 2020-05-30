using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpPropertyPart
    {
        public IAmAProperty Parent;
        public bool Encountered;
        public List<string> LineList = new List<string>();
        public List<string> BottomLineList = new List<string>();
        public List<Parameter> ParameterList = new List<Parameter>();
        public PropertyPartType PartType;
        public bool IsEmpty => (LineList.Where(x => x.Trim().Length > 0).ToList().Count
                              + BottomLineList.Where(x => x.Trim().Length > 0).ToList().Count) == 0;

        public CSharpPropertyPart(IAmAProperty parent, PropertyPartType propertyPartType)
        {
            Parent = parent;
            PartType = propertyPartType;
        }

        public string GenerateCode()
        {
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty)
                return Tools.Indent(Parent.Indent + 1) + $"{finalPartType};";

            var result = new StringBuilder();

            result.AppendLine($"{Tools.Indent(Parent.Indent + 1)}{finalPartType}");
            result.AppendLine($"{Tools.Indent(Parent.Indent + 1)}{{");
            foreach(var line in Massage.DetermineWhichLinesGetASemicolon(LineList))
            {
                result.AppendLine(Tools.Indent(Parent.Indent + 2) + line);
            }
            foreach(var line in Massage.DetermineWhichLinesGetASemicolon(BottomLineList))
            {
                result.AppendLine(Tools.Indent(Parent.Indent + 2) + line);
            }
            result.AppendLine($"{Tools.Indent(Parent.Indent + 1)}}}");

            return result.ToString();
        }
    }
}