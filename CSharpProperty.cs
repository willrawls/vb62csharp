using System;
using System.CodeDom.Compiler;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpProperty : AbstractCodeBlock, IAmAProperty
    {
        public CSharpPropertyPart Get;
        public CSharpPropertyPart Let;
        public CSharpPropertyPart Set;

        public string Comment { get; set; }
        public string Name { get; set; }
        public string Scope { get; set; }
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

        public void ConvertSourcePropertyParts(IAmAProperty sourceProperty)
        {
            var localSourceProperty = (Property)sourceProperty;
            CSharpPropertyPart targetPart;

            if (localSourceProperty.Direction == "Get")
                targetPart = Get;
            else if (localSourceProperty.Direction == "Let")
                targetPart = Let;
            else
                targetPart = Set;

            targetPart.ParameterList = localSourceProperty.Parameters;
            targetPart.Encountered = true;

            if (targetPart.LineList == null)
                targetPart.LineList = new CodeBlock(this);
            if (targetPart.BottomLineList == null)
                targetPart.BottomLineList = new CodeBlock(this);

            foreach (var originalLine in localSourceProperty.LineList.Children)
            {
                var line = originalLine.Line.Trim();
                if (!line.IsNotEmpty()) continue;

                ConvertSource.GetPropertyLine(line, out var translatedLine, out var placeAtBottom, localSourceProperty);
                if (translatedLine.IsNotEmpty())
                    targetPart.LineList.Children.Add(new CodeBlock(this, translatedLine));
                if (placeAtBottom.IsNotEmpty())
                    targetPart.BottomLineList.Children.Add(new CodeBlock(this, placeAtBottom));
                targetPart.Encountered = true;
            }
        }

        public override string GenerateCode()
        {
            var result = new StringBuilder();
            var firstIndentation = Tools.Indent(Indent);
            var secondIndentation = Tools.Indent(Indent + 1);

            // possible comment
            if (Comment.IsNotEmpty())
                if (!Comment.Contains("\n"))
                    result.AppendLine($"{firstIndentation}// {Comment.Substring(1)}");
                else
                {
                    Comment =
                        string.Join("\n",
                            Comment
                            .Replace("' ", "")
                            .Replace("\r", "")
                            .Split('\n')
                            .Select(x => firstIndentation + (x.EndsWith(";") ? x.Substring(0, x.Length - 1) : x)));
                    if (Comment.IsNotEmpty())
                    {
                        result.AppendLine();
                        result.AppendLine(firstIndentation + "/*");
                        result.AppendLine();
                        result.AppendLine(Comment);
                        result.AppendLine();
                        result.AppendLine(firstIndentation + "*/");
                        result.AppendLine();
                    }
                }

            var letSet = Set.Encountered ? Set : Let;
            var blockName = $"public {Type ?? "object"} {Name}";

            if (Get.Encountered && letSet.Encountered)
            {
                if (Get.IsEmpty && letSet.IsEmpty)
                {
                    result.AppendLine(secondIndentation + blockName + " { get; set; }");
                }
                else
                {
                    result.AppendLine(Get.GenerateCode());
                    result.AppendLine(letSet.GenerateCode());
                }
            }
            else if (Get.Encountered)
            {
                result.AppendLine(Get.GenerateCode());
                /*
                result.AppendLine(firstIndentation + $"public {Type} {Name}");
                result.AppendLine(firstIndentation + "{");
                if (Get.IsEmpty)
                    result.AppendLine(secondIndentation + "get; set; // Was get only");
                else
                    result.AppendLine(Get.GenerateCode());
                result.AppendLine(firstIndentation + "}");
                */
            }
            else if (letSet.Encountered)
            {
                if (letSet.IsEmpty)
                    result.AppendLine(secondIndentation + " { get; set; } // Was set only");
                else
                    result.AppendLine(letSet.GenerateCode());
            }

            var code =
                string.Join("\n",
                    result.ToString()
                        .Replace("\r", "")
                        .Split('\n')
                        .Where(x => x.Trim().IsNotEmpty())
                        .Select(x => x)
                );
            return code;
        }

        public string GenerateCode(CodeBlock parent)
        {
            Parent = parent;
            return GenerateCode();
        }
    }
}