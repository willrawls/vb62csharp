using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class Extensions
    {
        public static void ExamineAndAdjust(this IList<ICodeLine> lines)
        {
            throw new Exception("Start here");
        }

        public static string AddIndent(this string line, int indentLevelToAdd = 1)
        {
            if (indentLevelToAdd < 1)
                return line;

            var toAdd = indentLevelToAdd * 4;
            return new string(' ', toAdd) + line;
        }

        public static string Indent(this string originalLines, int indentLevel = 1)
        {
            if (indentLevel < 1)
                return originalLines;

            var result = new StringBuilder();
            var lines = originalLines.Replace("\r", "").Split('\n');
            var indentForLine = Tools.Indent(indentLevel);
            foreach (var line in lines)
            {
                result.AppendLine(indentForLine + line.Trim());
            }

            var code = result.ToString();
            return code;
        }

        public static bool IsNotEmpty<T>(this IList<T> target)
        {
            return target?.Count > 0;
        }

        public static string RemoveEmptyLines(this string target)
        {
            return string.Join("\n", target
                    .Replace("\r", "")
                    .Split('\n')
                    .Where(x => x.Trim().IsNotEmpty())
                    .Select(x => x)
            );
        }
    }
}