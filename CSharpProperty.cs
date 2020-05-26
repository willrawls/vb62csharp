using System;
using System.Collections.Generic;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpProperty : IAmAProperty
    {
        public string Direction;
        public string Scope { get; set; }

        public string Comment { get; set; }
        public string Name { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }

        public CSharpPropertyPart Get = new CSharpPropertyPart();
        public CSharpPropertyPart Set = new CSharpPropertyPart();

        public string GenerateTargetCode()
        {
            var result = new StringBuilder();

            // possible comment
            if(Comment.IsNotEmpty() && Comment != "'\r\n")
                result.AppendLine("// " + Comment + ";");

            if (Get.Encountered && Set.Encountered)
            {
                result.AppendLine("{");
                result.AppendLine(Get.GenerateCode());
                result.AppendLine(Set.GenerateCode());
                result.AppendLine("}");
            }
            else if (Get.Encountered)
            {
                result.AppendLine("{");
                result.AppendLine(Get.GenerateCode());
                result.AppendLine("}");
            }
            else if (Set.Encountered)
            {
                result.AppendLine("{");
                result.AppendLine(Set.GenerateCode());
                result.AppendLine("}");
            }
            else
            {
                throw new NotSupportedException();
            }

            /*
            result.AppendLine($"{ConvertCode.Indent4}public {Type} {Name}");
            if (LineList.Count == 0)
            {
                result.AppendLine($"{ConvertCode.Indent6}{{ get; ");
            }
            else
            {
                result.AppendLine($"{ConvertCode.Indent6}{{ get {{");
            }

            // lines
            var atBottom = new List<string>();
            foreach (var originalLine in LineList)
            {
                var line = originalLine.Trim();
                if (line.Length > 0)
                {
                    Tools.ConvertLineOfCode(line, out var convertedLineOfCode, out var placeAtBottom);
                    if (convertedLineOfCode.IsNotEmpty())
                        convertedLineOfCode = ConvertCode.Indent6 + convertedLineOfCode + ";";
                    result.AppendLine(convertedLineOfCode);
                    if (placeAtBottom.IsNotEmpty())
                        atBottom.Add(placeAtBottom);
                }
                else
                {
                    result.AppendLine();
                }
            }

            foreach (var line in atBottom)
                result.AppendLine(line);

            result.Append(ConvertCode.Indent4 + "}\r\n");
            */

            return result.ToString();
        }

        public void Convert(IAmAProperty sourceProperty)
        {
            var localSourceProperty = (Property) sourceProperty;
            var result = new StringBuilder();
            CSharpPropertyPart targetPart;

            if(localSourceProperty.Direction == "Get")
            {
                targetPart = Get;
                targetPart.ParameterList = localSourceProperty.Parameters;
                targetPart.Encountered = true;
                
            }
            else
            {
                targetPart = Set;
            }

            foreach (var originalLine in localSourceProperty.LineList)
            {
                var line = originalLine.Trim();
                if(line.IsNotEmpty())
                { 
                    Tools.ConvertLineOfCode(line, out var translatedLine, out var placeAtBottom, localSourceProperty);
                    if(translatedLine.IsNotEmpty())
                        targetPart.LineList.Add(translatedLine);
                    if (placeAtBottom.IsNotEmpty())
                        targetPart.BottomLineList.Add(placeAtBottom);
                    targetPart.Encountered = true;
                }
            }

            
            throw new NotImplementedException();
        }
    }
}