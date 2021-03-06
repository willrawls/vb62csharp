using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public class CSharpPropertyPart : AbstractBlock, IGenerate, ICodeLine
    {
        public CodeLineList LineListAfter;

        public bool Encountered;

        public List<Parameter> ParameterList;
        public PropertyPartType PartType;

        public CSharpPropertyPart(ICodeLine parent, PropertyPartType propertyPartType) : base(parent, null)
        {
            Parent = parent;
            PartType = propertyPartType;
            ParameterList = new List<Parameter>();
            LineListAfter = new CodeLineList();
            ResetIndentsRecursively = CSharpPropertyPartResetIndentsRecursively;
        }

        private int CSharpPropertyPartResetIndentsRecursively(int indentLevel)
        {
            Indent = indentLevel;
            foreach (var child in Children)
            {
                child.ResetIndent(indentLevel + 1);
            }
            return indentLevel;
        }

        public override string GenerateCode()
        {
            // ResetIndent(indentLevel);

            var finalPartType = PartType.ToString().ToLower().Replace("let", "set");

            if (IsEmpty())
                return Indentation + $"{finalPartType};";

            var result = new StringBuilder();
            result.AppendLine(Indentation + $"{finalPartType}");

            if(Before.IsNotEmpty())
            {
                result.AppendLine(Indentation + Before);
            }

            if(Line.IsNotEmpty())
            {
                result.AppendLine(SecondIndentation + Line);
                if (Before.IsNotEmpty())
                    result.AppendLine(SecondIndentation + Before);
            }


            foreach (var codeLine in Children.Where(block => block.Line.IsNotEmpty()))
            {
                var generatedCode = codeLine.GenerateCode();
                string massagedLine;

                if (Line.IsEmpty())
                    massagedLine = generatedCode;
                else
                    massagedLine = generatedCode.AddIndent();
                result.Append(massagedLine);
            }

            foreach (var blockAfter in LineListAfter.Where(line => line.IsNotEmpty()))
            {
                // blockAfter.ResetIndent(indentLevel + 1);
                var generatedCode = blockAfter.GenerateCode();
                string blockLine;
                if (Line.IsEmpty())
                    blockLine = generatedCode;
                else
                    blockLine = generatedCode.AddIndent();
                result.Append(blockLine);
            }

            if (Line.IsNotEmpty())
            {
                result.AppendLine(SecondIndentation + After);
            }

            if(After.IsNotEmpty())
            {
                result.AppendLine(Indentation + After);
            }

            var code = result.ToString();
            return code;
        }

        public void ExamineAndAdjust()
        {
            Line = Line.ExamineAndAdjustLine(CurrentlyInArea.Property);
            Children.ExamineAndAdjust(CurrentlyInArea.Property);
            LineListAfter.ExamineAndAdjust(CurrentlyInArea.Property);
        }
    }
}