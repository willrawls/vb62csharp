using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class Massage
    {
        private static List<string> LineContainsAnyOfTheseThenShouldHaveASemiColon =
            new List<string>
            {
                "{", "}",
                "foreach",
                "while",
                "if(", 
                "if (", 
                "else",
            };

        private static List<string> LineDoesNotContainAnyOfThese = new List<string>
        {
            "//",
        };

        /// <summary>
        ///     When line starts with X and ends with Y:
        ///     Remove X and Y
        ///     Add Z as a line above the beginning
        ///     Add A as a line below the end
        /// </summary>
        public static List<XReplace> AboveAndBelowReplacements { get; set; } =
            new List<XReplace>
            {
            };

        /// <summary>
        ///     When line contains X:
        ///     Replace with Y
        /// </summary>
        public static List<XReplace> BlanketReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace("Exit Property", "return; // ???"),
                new XReplace("Exit Function", "return; // ???"),
                new XReplace("Exit Sub", "return;"),
                new XReplace(" In ", " in "),
                new XReplace(";;", ";"),
                new XReplace("{;", "{"),
                new XReplace("; = ", " = "),
                new XReplace("{ get; set; }", ";"),
                new XReplace("Static", "static"),
                new XReplace("String", "string"),
                new XReplace("Collection", "Dictionary<string,string>"),
                new XReplace("System.Dictionary<string,string>", "System.Collection"),
                new XReplace("using System.Dictionary<string,string>()s", "using System.Collections"),
                new XReplace("True", "true"),
                new XReplace("False", "false"),
                new XReplace("private ", "public "), // Because I believe everything should be public
                new XReplace("Err.Clear", string.Empty),
                new XReplace("Err.Number", "ex"),
                new XReplace("Me.", "this."),
                new XReplace("/*;", "/*"),
                new XReplace("*/;", "*/"),
                new XReplace("); ); )", ")"),
                new XReplace("); )", ")"),
                new XReplace("Resume ", "goto "),
                new XReplace("Long ", "long "),
                new XReplace("Clear ", "this.Clear()"),
                new XReplace("()()", "()"),
                new XReplace("Add(", "Add( "),
                new XReplace("Next ", "} // "),
                new XReplace(" & ", " + "),
                new XReplace("Integer", "int"),
                new XReplace(".[", "["),
                new XReplace(" As ", " /*As*/ "),
            };

        /// <summary>
        ///     When line ends with X:
        ///     Replace X with Y
        /// </summary>
        public static List<XReplace> EndsWithReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace(";;", ";"),
                new XReplace(":;", "~~~"),
            };

        /// <summary>
        ///     When line starts with X and ends with Y:
        ///     Replace Z with A
        /// </summary>
        public static List<XReplace> StartsAndEndsWithReplacements { get; set; } =
            new List<XReplace>()
            {
                new XReplace("Set", "Nothing", "Set", ""),
            };

        /// <summary>
        ///     When line starts with X, Replace X with Y, Replace Z with A
        /// </summary>
        public static List<XReplace> StartsWithReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace("' ", "// "),
                new XReplace("foreach( var ", "foreach(var ", " )", " ) {"),
                new XReplace("On Error GoTo ", "// TODO: Rewrite try/catch and/or goto. "),
            };

        /// <summary>
        ///     When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static List<XReplace> WhenStartsWithReplaceOtherReplacements { get; set; } =
            new List<XReplace>()
            {
                new XReplace("foreach(", null, null, " )"),
            };

        /*
        public static string FindCodeBetweenBraces(this string target)
        {
            var input = target;
            var regex = new Regex("{((?>[^{}]+|{(?<c>)|}(?<-c>))*(?(c)(?!)))", RegexOptions.Singleline);
            var matches = regex.Matches(input);
            if(matches?.Count > 0)
            {
                var result = matches[0].Groups[1].Value;
                return result;
            }

            return "";
        }
        */

        public static Block AsBlock(this string target, ICodeLine parent)
        {
            var result = Quick.Block(parent, null);
            var inner = target;
            while (inner.FindCodeBetweenBraces(
                out var before,
                out var insideBraces,
                out var after,
                out var index))
            {
                foreach (var lineBefore in before.Lines())
                    result.Children.Add(Quick.Line(result, lineBefore));

                var innerBlock = Quick.Block(result, null);
                foreach(var innerLine in insideBraces.Lines().Where(x => x.IsNotEmpty()))
                    innerBlock.Children.Add(Quick.Line(innerLine));

                foreach (var lineAfter in after.Lines().Where(x => x.IsNotEmpty()))
                    result.Children.Add(Quick.Line(result, lineAfter));

                result = innerBlock;
            }

            return result;
        }

        public static string Regexify(this string target)
        {
            var result = target
                .Replace("{", "\\{")
                .Replace("}", "\\}")
                .Replace("[", "\\[")
                .Replace("]", "\\]")
                .Replace("(", "\\(")
                .Replace(")", "\\)")
                .Replace(".", "\\.")
                .Replace("*", "\\*");
            return result;
        }

        /// <summary>
        ///     When line starts with X and ends with Y:
        ///     Remove X and Y,
        ///     Add Z as a line above the beginning,
        ///     Add A as a line below the end
        /// </summary>
        public static string AboveAndBelowWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in AboveAndBelowReplacements)
            {
                while (lineOfCode.ToLower().StartsWith(entry.X.ToLower()) &&
                       lineOfCode.ToLower().EndsWith(entry.Y.ToLower()))
                {
                    lineOfCode = lineOfCode
                        .Replace(entry.X, "")
                        .Replace(entry.Y, "");

                    lineOfCode = entry.Z + Environment.NewLine + lineOfCode + Environment.NewLine +
                                 entry.A;
                }
            }

            return lineOfCode;
        }

        /// <summary>
        ///     When line contains X, replace with Y
        /// </summary>
        /// <param name="originalLineOfCode"></param>
        /// <returns></returns>
        public static string BlanketReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in BlanketReplacements)
            {
                var iterations = 0;
                while (++iterations < 100 && lineOfCode.ToLower().Contains(entry.X.ToLower()))
                    lineOfCode = lineOfCode.Replace(entry.X, entry.Y);
            }

            return lineOfCode;
        }

        /// <summary>
        ///     Special case replacements
        /// </summary>
        /// <param name="lineOfCode"></param>
        /// <returns></returns>
        public static string CleanupTranslatedLineOfCode(string originalLineOfCode)
        {
            var lineOfCode = originalLineOfCode.Replace("//SOB//", " {");
            if (lineOfCode.TokenCount() == 4 && lineOfCode.TokenAt(3) == "As")
            {
                lineOfCode = $"{lineOfCode.TokenAt(4)} {lineOfCode.TokenAt(2)};";
                if (lineOfCode.Contains("static "))
                    lineOfCode = "static " + lineOfCode;
            }

            if (lineOfCode.Contains("foreach(") && lineOfCode.EndsWith(");"))
                lineOfCode = lineOfCode.FirstToken(");") + ")";

            if (lineOfCode.StartsWith("Next "))
                lineOfCode = "}\r\n";

            if (lineOfCode == "'" || lineOfCode == ";")
                lineOfCode = "";

            lineOfCode = lineOfCode
                .Replace("\r", "")
                .Replace("\n\n", "\n")
                .Replace("\n\n", "\n")
                .Replace("\n\n", "\n");

            return lineOfCode;
        }

        /// <summary>
        ///     Adds semicolon to end of appropriate lines
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static string DetermineIfLineGetsASemicolon(string line, string nextLine)
        {

            if (line.IsEmpty())
                return string.Empty;

            if (line.Trim().EndsWith(";"))
                return line;

            if (LineContainsAnyOfTheseThenShouldHaveASemiColon
                    .Any(x => line.Contains(x)) == false
                && LineDoesNotContainAnyOfThese.Any(x => line.Contains(x) == false && !line.EndsWith(":")))
            {
                if(nextLine.IsEmpty() || !nextLine.Contains("{"))
                     return line + ";";
            }

            if(line.Trim() == ";")
                line = "";
            return line;
        }

        /// <summary>
        ///     Adds semicolon to end of appropriate lines
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static IList<string> DetermineWhichLinesGetASemicolon(List<string> lines)
        {
            if (lines.IsEmpty())
                return lines;

            for (var i = 0; i < lines.Count; i++)
            {
                lines[i] = DetermineIfLineGetsASemicolon(lines[i], i+1 < lines.Count ? lines[i+1] : null);
            }

            return lines;
        }

        /// <summary>
        ///     When line ends with X
        ///     Replace X with Z
        /// </summary>
        /// <param name="originalLineOfCode"></param>
        /// <returns></returns>
        public static string EndsWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in EndsWithReplacements)
            {
                while (lineOfCode.ToLower().EndsWith(entry.X.ToLower()))
                    lineOfCode = lineOfCode.Replace(entry.X.ToLower(), entry.Z);
            }

            return lineOfCode;
        }

        public static string AllLinesNow(string translatedLine)
        {
            if (translatedLine.IsEmpty())
                return string.Empty;

            var result = new StringBuilder();
            var lines = translatedLine
                .Replace("\r", string.Empty)
                .Split('\n').ToList();

            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var nextLine = i < lines.Count - 1 ? lines[i + 1] : null;
                result.AppendLine(Transform(line, nextLine));
            }

            var code = result.ToString();
            return code;
        }

        /// <summary>
        ///     Run all blanket and specialized replacements now
        /// </summary>
        /// <param name="translatedLine"></param>
        /// <returns></returns>
        public static string Transform(this string originalLine, string nextLine = null)
        {
            if (originalLine.IsEmpty())
                return "";

            if (originalLine.Trim() == "")
                return originalLine;

            if (originalLine.Trim().StartsWith("//"))
                return originalLine;

            var translatedLine = originalLine;
            translatedLine = translatedLine.HandleWith();

            translatedLine = BlanketReplaceNow(translatedLine);
            translatedLine = StartsWithReplaceNow(translatedLine);
            translatedLine = EndsWithReplaceNow(translatedLine);

            translatedLine = StartsAndEndsWithReplaceNow(translatedLine);
            translatedLine = AboveAndBelowWithReplaceNow(translatedLine);

            translatedLine = WhenStartsWithReplaceOtherAndAppendNow(translatedLine);

            translatedLine = RegExReplacements.Shuffle(translatedLine);

            translatedLine = CleanupTranslatedLineOfCode(translatedLine);
            translatedLine = DetermineIfLineGetsASemicolon(translatedLine, nextLine);

            if (translatedLine.Trim() == ";")
                translatedLine = string.Empty;
            
            return translatedLine;
        }

        /// <summary>
        ///     When line starts with X and ends with Y
        ///     Replace Z with A
        /// </summary>
        public static string StartsAndEndsWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in StartsAndEndsWithReplacements)
                if (lineOfCode.ToLower().StartsWith(entry.X.ToLower()) &&
                    lineOfCode.ToLower().EndsWith(entry.Y.ToLower()))
                    lineOfCode = lineOfCode.Replace(entry.Z, entry.A);

            return lineOfCode;
        }

        /// <summary>
        ///     When line starts with X, Replace X with Y, Replace Z with A
        /// </summary>
        public static string StartsWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in StartsWithReplacements)
            {
                var iterations = 0;
                while (++iterations < 100 && lineOfCode.ToLower().StartsWith(entry.X.ToLower()))
                {
                    if (entry.Y.IsNotEmpty())
                        lineOfCode = lineOfCode.Replace(entry.X, entry.Y);
                    if (entry.Z.IsNotEmpty() && entry.A.IsNotEmpty())
                        lineOfCode = lineOfCode.Replace(entry.Z, entry.A);
                }

                if (iterations > 98)
                    lineOfCode += " // TODO: Check: Too many iterations during StartsWithReplaceNow";
                //throw new Exception("Too many tries in StartsWithReplaceNow");
            }

            return lineOfCode;
        }

        /// <summary>
        ///     When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static string WhenStartsWithReplaceOtherAndAppendNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in WhenStartsWithReplaceOtherReplacements)
                if (entry.X.IsNotEmpty() && lineOfCode.ToLower().StartsWith(entry.X.ToLower()))
                {
                    if (entry.Y.IsNotEmpty() && entry.Z.IsNotEmpty())
                        lineOfCode = lineOfCode.Replace(entry.Y, entry.Z);
                    if (entry.A.IsNotEmpty())
                        lineOfCode += entry.A ?? string.Empty;
                }

            return lineOfCode;
        }

        public static bool FindCodeBetweenBraces(this string target, out string before, out string insideTheMostOuterBraces, out string after, out int indexOfMostInner)
        {
            var input = target;
            var regex = new Regex("{((?>[^{}]+|{(?<c>)|}(?<-c>))*(?(c)(?!)))", RegexOptions.Singleline);
            var splits = regex.Split(input).Where(s => s.Length > 0).ToArray();
            
            if (splits.Length > 0)
            {
                before = splits[0];
                insideTheMostOuterBraces = splits[1];
                after = splits[2];
                indexOfMostInner = target.IndexOf(insideTheMostOuterBraces, StringComparison.InvariantCultureIgnoreCase);
                return true;
            }

            before = null;
            insideTheMostOuterBraces = target;
            after = null;
            indexOfMostInner = 0;
            return false;
        }
    }
}