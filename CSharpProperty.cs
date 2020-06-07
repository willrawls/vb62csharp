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

            targetPart.Children.Add(new Block(this));
            targetPart.LinesAfter = new Block(this);

            foreach (var originalLine in localSourceProperty.Block.Children)
            {
                var line = originalLine.Line.Trim();
                if (!line.IsNotEmpty()) continue;

                ConvertSource.GetPropertyLine(line, out var translatedLine, out var placeAtBottom, localSourceProperty);
                if (translatedLine.IsNotEmpty())
                    targetPart.Children.Add(new Block(this, translatedLine));
                if (placeAtBottom.IsNotEmpty())
                    targetPart.LinesAfter.Children.Add(new Block(this, placeAtBottom));
                targetPart.Encountered = true;
            }
        }

        public override string GenerateCode()
        {
            var result = new StringBuilder();

            // possible comment
            if (Comment.IsNotEmpty())
                if (!Comment.Contains("\n"))
                    result.AppendLine($"{Indentation}// {Comment.Substring(1)}");
                else
                {
                    Comment =
                        string.Join("\n",
                            Comment
                            .Replace("' ", "")
                            .Replace("\r", "")
                            .Split('\n')
                            .Select(x => SecondIndentation + (x.EndsWith(";") ? x.Substring(0, x.Length - 1) : x)));
                    if (Comment.IsNotEmpty())
                    {
                        result.AppendLine();
                        result.AppendLine(Indentation + "/*");
                        result.AppendLine(Comment);
                        result.AppendLine(Indentation + "*/");
                        result.AppendLine();
                    }
                }

            var letSet = Set.Encountered ? Set : Let;
            var propertyHeader = $"public {Type ?? "object"} {Name}";

            if(Line.IsNotEmpty())
            {
                result.AppendLine(Indentation + Line);
            }

            if (Get.IsEmpty() && letSet.IsEmpty())
            {
                result.AppendLine(Indentation + propertyHeader + " { get; set; }");
            }
            else
            {
                result.AppendLine(Indentation + propertyHeader);
                result.AppendLine(Indentation + Before);

                if (Get.Encountered)
                {
                    result.AppendLine(Get.GenerateCode());
                }

                if(letSet.Encountered)
                {
                    if (letSet.IsEmpty())
                        result.AppendLine(SecondIndentation + "{ get; set; } // Was set only");
                    else
                        result.AppendLine(letSet.GenerateCode());
                }
                result.AppendLine(Indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public string GenerateCode(Block parent)
        {
            Parent = parent;
            return GenerateCode();
        }
    }
}