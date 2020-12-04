using System;
using System.Collections.Generic;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace MetX.VB6ToCSharp.CSharp
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
        public Module TargetModule { get; set; }

        public CSharpProperty(ICodeLine parent) : base(parent)
        {
            ResetIndentsRecursively = CSharpPropertyResetIndentsRecursively;
            Get = new CSharpPropertyPart(this, PropertyPartType.Get);
            Set = new CSharpPropertyPart(this, PropertyPartType.Set);
            Let = new CSharpPropertyPart(this, PropertyPartType.Let);

        }

        private int CSharpPropertyResetIndentsRecursively(int indentLevel)
        {
            Get.ResetIndent(indentLevel + 1);
            Set.ResetIndent(indentLevel + 1);
            Let.ResetIndent(indentLevel + 1);
            return indentLevel;
        }

        public void ConvertParts(IAmAProperty sourceProperty)
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

            var block = new Block(this);
            targetPart.Children.Add(block);

            for (var i = 0; i < localSourceProperty.Block.Children.Count; i++)
            {
                var originalLine = localSourceProperty.Block.Children[i];
                var nextLine = i < localSourceProperty.Block.Children.Count - 1 
                    ? localSourceProperty.Block.Children[i+1]
                    : null;
                var line = originalLine.Line.Trim();
                if (line.IsEmpty()) continue;

                ConvertSource.GetPropertyLine(
                    line, 
                    nextLine?.Line, 
                    out var translatedLine,
                    out var placeAtBottom, 
                    localSourceProperty);

                if (translatedLine.IsNotEmpty())
                    targetPart.Children.Add(new LineOfCode(this, translatedLine));
                if (placeAtBottom.IsNotEmpty())
                {
                    var linesAtBottom = placeAtBottom.Lines(StringSplitOptions.RemoveEmptyEntries);
                    foreach(var bottomLine in linesAtBottom)
                        targetPart.LineListAfter.Add(new LineOfCode(this, bottomLine));
                }
                targetPart.ExamineAndAdjust();
                targetPart.ExamineAndAdjust();

                targetPart.Encountered = true;
            }
        }

        public static string DetermineBackingVariable(CSharpProperty target)
        {
            if (target.Get == null
                || !target.Get.Encountered
                || target.Get.Children == null
                || target.Get.Children?.Count == 0)
                return "object /*unknown(dbv1)*/";

            foreach (var child in target
                .Get
                .Children
                .Where(x => x
                    .Line?.Contains("return") == true))
            {
                var possible = child.Line
                    .TokenBetween("return", ";")
                    .Trim();
                if (possible.IsNotEmpty() && !possible.Contains(" "))
                {
                    return possible;
                }
            }

            return null;
        }

        public override string GenerateCode()
        {
            ResetIndent(Indent); // Will set recursively

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
                        result.AppendLine(Comment);
                        result.AppendLine();
                    }
                }

            var letSet = Set.Encountered ? Set : Let;

            if (string.IsNullOrEmpty(Type))
            {
                var backingVariable = CSharpProperty.DetermineBackingVariable(this);
                if (TargetModule.VariableList.Contains(backingVariable))
                {
                    Type = TargetModule.VariableList[backingVariable].Type;
                }
                else
                {
                    // TODO
                }
            }

            var propertyHeader = $"public {Type ?? "unknownTypeFor" + Name} {Name}";

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
                    var getCode = Get.GenerateCode();
                    result.AppendLine(getCode);
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
    }
}