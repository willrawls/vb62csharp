using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class Extensions
    {
        public static string CSharpTypeEquivalent(this string target)
        {
            if (target.IsEmpty())
                return $"UnknownCSharpEquivalent({target})";

            switch (target.ToLower().Trim())
            {
                case "string": return "string";
                case "collection": return "Dictionary<string,string>";
                default: return target;
            }
        }
        public static string Deflate(this string target)
        {
            if (target.IsEmpty())
                return "";
            
            target = target.Trim();
            var tokens = target.AllTokens("\"");
            var result = "";

            for (var index = 0; index < tokens.Count; index++)
            {
                if (index % 2 == 0)
                {
                    var token = tokens[index];
                    while (token.Contains("  "))
                        token = token.Replace("  ", " ");
                    result += token;
                }
                else
                    result += tokens[index];

                if (index < tokens.Count - 1)
                    result += "\"";
            }

            return result;
        }


        public static string PutABeforeBOnce(this string target, string a, string b)
        {
            if (target.IsEmpty() || !target.Contains(a) || !target.Contains(b))
                return target ?? "";

            if (string.Compare(
                target, 
                b + a,
                CultureInfo.CurrentCulture, CompareOptions.IgnoreCase ) == 0)
                return a + b;

            var result = "";

            var indexOfB = target.IndexOf(a, StringComparison.InvariantCultureIgnoreCase);
            var indexOfA = target.IndexOf(b, StringComparison.InvariantCultureIgnoreCase);

            if (indexOfB < indexOfA)
                return target;

            /*
            if (indexOfB > 0)
                result += target.Substring(0, indexOfB);
            result += a;
            */

            // -----Xxx.....Yy=====
            // 0for5 + 13for2 + 8for5 + 5for3 + 15for5
            // indexOfA = 13    length of A = 2  A = Yy
            // indexOfB = 5     length of B = 3  B = Xxx
            // -----Yy.....Xxx=====
            // ----- @ 0    for 5   segment 1
            // ..... @ 8    for 5   segment 3   length = 13 - (5 + 3)
            // ===== @ 15   for 5   segment 5
            //


            var segment1 = target.Substring(0, indexOfA);
            var segment2 = target.Substring(segment1.Length, b.Length);
            
            var segment3Length  = indexOfB - (indexOfA + b.Length);
            var segment3 = segment3Length > 0
                ? target.Substring(segment1.Length + segment2.Length, segment3Length)
                : "";

            var segment4 = target.Substring(segment1.Length + segment2.Length + segment3.Length, a.Length);
            
            var segment5 = target.Substring(segment1.Length + segment2.Length + segment3.Length + segment4.Length);

            result = $"{segment1}{segment4}{segment3}{segment2}{segment5}";
            /*
            if (result.Length < indexOfA)
                result += target.Substring(result.Length, indexOfA - indexOfB + b.Length);
            result += b;

            if(result.Length < target.Length)
                result += target.Substring(indexOfA + a.Length);
                */

            return result;
        }

        public static string ExamineAndAdjustLine(this string target, CurrentlyInArea inArea)
        {
            if (target.IsEmpty())
                return "";

            switch (inArea)
            {
                case CurrentlyInArea.Class:
                    target = target.PutABeforeBOnce("static", "string");
                    target = target.PutABeforeBOnce("static", "int");
                    target = target.PutABeforeBOnce("static", "long");
                    target = target.PutABeforeBOnce("static", "Collection");
                    break;
                case CurrentlyInArea.Procedure:
                    target = target.Replace("static", "").Deflate();
                    break;
                case CurrentlyInArea.Property:
                    target = target.Replace("static", "").Deflate();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(inArea), inArea, null);
            }

            return target;
        }

        public static void ExamineAndAdjust(this CodeLineList targets, CurrentlyInArea inArea)
        {
            if (targets.IsEmpty())
                return;

            foreach (var target in targets.Where(x => x is LineOfCode).Cast<LineOfCode>())
            {
                target.Line = target.Line.ExamineAndAdjustLine(inArea);
            }

            foreach (var target in targets.Where(x => x is AbstractBlock).Cast<AbstractBlock>())
            {
                target.Line = target.Line.ExamineAndAdjustLine(inArea);
                target.Children.ExamineAndAdjust(inArea);
            }
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