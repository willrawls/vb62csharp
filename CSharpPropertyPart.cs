using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
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
            var firstIndent = Tools.Indent(Parent.Indent + 1);
            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty)
                return firstIndent + $"{finalPartType};";

            var result = new StringBuilder();

            result.AppendLine(Tools.Blockify($"{firstIndent}{finalPartType}", Parent.Indent + 1, builder =>
            {
                var blockBuilder = new StringBuilder();
                foreach (var line in Massage
                    .DetermineWhichLinesGetASemicolon(LineList)
                    .Where(x => x.Trim() != "")
                    .Select(x => Massage.Now(x)))
                    blockBuilder.AppendLine(line);

                foreach (var line in Massage
                    .DetermineWhichLinesGetASemicolon(BottomLineList)
                    .Where(x => x.Trim() != "")) 
                    blockBuilder.AppendLine(line.Trim());

                var blockCode = blockBuilder.ToString();
                return blockCode;
            }));

            var code = result.ToString();
            return code;
        }
    }
}