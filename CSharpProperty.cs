using System;
using System.CodeDom.Compiler;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpProperty : IAmAProperty
    {
        public CSharpPropertyPart Get;
        public CSharpPropertyPart Let;
        public CSharpPropertyPart Set;

        public string Comment { get; set; }
        public string Name { get; set; }
        public string Scope { get; set; }
        public int Indent { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }

        public CSharpProperty(int parentIndent)
        {
            Get = new CSharpPropertyPart(this, PropertyPartType.Get);
            Set = new CSharpPropertyPart(this, PropertyPartType.Set);
            Let = new CSharpPropertyPart(this, PropertyPartType.Let);
            Indent = parentIndent + 1;

        }

        public void ParsePropertyParts(IAmAProperty sourceProperty)
        {
            var localSourceProperty = (Property) sourceProperty;
            CSharpPropertyPart targetPart;

            if (localSourceProperty.Direction == "Get")
                targetPart = Get;
            else if (localSourceProperty.Direction == "Let")
                targetPart = Let;
            else
                targetPart = Set;

            targetPart.ParameterList = localSourceProperty.Parameters;
            targetPart.Encountered = true;

            foreach (var originalLine in localSourceProperty.LineList)
            {
                var line = originalLine.Trim();
                if (line.IsNotEmpty())
                {
                    Tools.ConvertLineOfCode(line, out var translatedLine, out var placeAtBottom, localSourceProperty);
                    if (translatedLine.IsNotEmpty())
                        targetPart.LineList.Add(translatedLine);
                    if (placeAtBottom.IsNotEmpty())
                        targetPart.BottomLineList.Add(placeAtBottom);
                    targetPart.Encountered = true;
                }
            }
        }

        public string GenerateTargetCode()
        {
            var result = new StringBuilder();

            // possible comment
            if (Comment.IsNotEmpty() && Comment != "'\r\n")
                result.AppendLine(Tools.Indent(Indent) + "// " + Comment + ";");

            var letSet = Set.Encountered ? Set : Let;

            if (Get.Encountered && letSet.Encountered)
            {
                result.AppendLine(Tools.Indent(Indent) + $"public {Type ?? "object"} {Name}");
                if (Get.IsEmpty && letSet.IsEmpty)
                {
                    result.AppendLine(Tools.Indent(Indent + 1) + "{ get; set; }");
                }
                else
                {
                    result.AppendLine(Tools.Indent(Indent) + "{");
                    result.AppendLine(Get.GenerateCode());
                    result.AppendLine(letSet.GenerateCode());
                    result.AppendLine(Tools.Indent(Indent) + "}");
                }
            }
            else if (Get.Encountered)
            {
                result.AppendLine(Tools.Indent(Indent) + $"public {Type} {Name}");
                result.AppendLine(Tools.Indent(Indent) + "{");
                if(Get.IsEmpty)
                    result.AppendLine(Tools.Indent(Indent + 1) + "get; set; // Was get only");
                else
                result.AppendLine(Get.GenerateCode());

                result.AppendLine(Tools.Indent(Indent) + "}");
            }
            else if (letSet.Encountered)
            {
                result.AppendLine(Tools.Indent(Indent) + "{");
                if(letSet.IsEmpty)
                    result.AppendLine(Tools.Indent(Indent + 1) + " get; set; // Was set only");
                else
                    result.AppendLine(Tools.Indent(Indent) + letSet.GenerateCode());
                result.AppendLine(Tools.Indent(Indent) + "}");
            }

            return result.ToString();
        }
    }
}