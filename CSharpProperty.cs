using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public class CSharpProperty : AbstractBlock, IAmAProperty
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

        public CSharpProperty(ICodeLine parent) : base(parent, null)
        {
            Get = new CSharpPropertyPart(this, PropertyPartType.Get);
            Set = new CSharpPropertyPart(this, PropertyPartType.Set);
            Let = new CSharpPropertyPart(this, PropertyPartType.Let);
            Indent = parent.Indent + 1;
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

            if (targetPart.Children == null)
                targetPart.Children = new List<ICodeLine> { new Block(this) };
            if (targetPart.BlockAtBottom == null)
                targetPart.BlockAtBottom = new Block(this);

            foreach (var originalLine in localSourceProperty.Block.Children)
            {
                var line = originalLine.Line.Trim();
                if (!line.IsNotEmpty()) continue;

                ConvertSource.GetPropertyLine(line, out var translatedLine, out var placeAtBottom, localSourceProperty);
                if (translatedLine.IsNotEmpty())
                    targetPart.Children.Add(new Block(this, translatedLine));
                if (placeAtBottom.IsNotEmpty())
                    targetPart.BlockAtBottom.Children.Add(new Block(this, placeAtBottom));
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
                            .Select(x => secondIndentation + (x.EndsWith(";") ? x.Substring(0, x.Length - 1) : x)));
                    if (Comment.IsNotEmpty())
                    {
                        result.AppendLine();
                        result.AppendLine(firstIndentation + "/*");
                        result.AppendLine(Comment);
                        result.AppendLine(firstIndentation + "*/");
                        result.AppendLine();
                    }
                }

            var letSet = Set.Encountered ? Set : Let;
            var blockName = $"public {Type ?? "object"} {Name}";

            if (Get.Encountered && letSet.Encountered)
            {
                if (Get.IsEmpty() && letSet.IsEmpty())
                {
                    result.AppendLine(firstIndentation + blockName + " { get; set; }");
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
                if (letSet.IsEmpty())
                    result.AppendLine(secondIndentation + " { get; set; } // Was set only");
                else
                    result.AppendLine(letSet.GenerateCode());
            }

            var code = result.ToString().RemoveEmptyLines();
            return code;
        }

        public string GenerateCode(Block parent)
        {
            Parent = parent;
            return GenerateCode();
        }
    }
}