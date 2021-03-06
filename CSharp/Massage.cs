using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Module = MetX.VB6ToCSharp.VB6.Module;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class Massage
    {
        private static readonly List<string> LineContainsAnyOfTheseThenShouldHaveASemiColon =
            new List<string>
            {
                "{", "}",
                "foreach",
                "while",
                "if(",
                "if (",
                "else"
            };

        private static readonly List<string> LineDoesNotContainAnyOfThese = new List<string>
        {
            "//"
        };

        /// <summary>
        ///     When line starts with X and ends with Y:
        ///     Remove X and Y
        ///     Add Z as a line above the beginning
        ///     Add A as a line below the end
        /// </summary>
        public static List<XReplace> AboveAndBelowReplacements { get; set; } =
            new List<XReplace>();

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
                new XReplace("using System.Dictionary<string,string>()s",
                    "using System.Collections"),
                new XReplace("True", "true"),
                new XReplace("False", "false"),
                new XReplace("private ",
                    "public "), // Because I believe everything should be public
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
                //new XReplace("Add(", "Add( "),
                new XReplace("Next ", "} // "),
                new XReplace(" & ", " + "),
                new XReplace("Integer", "int"),
                new XReplace(".[", "["),
                new XReplace("Is Null", "== null"),
                new XReplace("Is null", "== null"),
                new XReplace("is null", "== null"),
                new XReplace("Loop;", "}"),
                new XReplace("Loop", "}"),
                new XReplace("void Class_Initialize", "CurrentModuleName"),
                new XReplace("void Class_Terminate", "~CurrentModuleName"),
                new XReplace("Class_Initialize", "CurrentModuleName"),
                new XReplace("Class_Terminate", "~CurrentModuleName"),
                new XReplace("Attribute ", "// Attribute "),
            };

        /// <summary>
        ///     When line ends with X:
        ///     Replace X with Y
        /// </summary>
        public static List<XReplace> EndsWithReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace(";;", ";"),
                new XReplace(":;", "~~~")
            };

        /// <summary>
        ///     When line starts with X and ends with Y:
        ///     Replace Z with A
        /// </summary>
        public static List<XReplace> StartsAndEndsWithReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace("Set", "Nothing", "Set", "")
            };

        /// <summary>
        ///     When line starts with X, Replace X with Y, Replace Z with A
        /// </summary>
        public static List<XReplace> StartsWithReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace("' ", "// "),
                new XReplace("foreach( var ", "foreach(var ", " )", " ) {"),
                new XReplace("On Error GoTo ", "// TODO: Rewrite try/catch and/or goto. ")
            };

        /// <summary>
        ///     When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static List<XReplace> WhenStartsWithReplaceOtherReplacements { get; set; } =
            new List<XReplace>
            {
                new XReplace("foreach(", null, null, " )")
            };


        public static bool MissingAny(this string target, string mustHave1, string mustHave2)
        {
            return target.MissingAny(new[] {mustHave1, mustHave2});
        }

        public static bool MissingAny(this string target, string[] mustHaves)
        {
            return !mustHaves.All(target.Contains);
        }

        public static IList<string> Trimmed(this IList<string> target)
        {
            if (target.IsEmpty())
                return target;

            var result = target.MakeACopy();
            for (var i = 0; i < target.Count; i++)
                if (target[i].IsNotEmpty())
                    target[i] = target[i].Trim();
            return target;
        }

        public static ICodeLine AsBlock(this string target, ICodeLine parent, bool noOuterBraces)
        {
            if (parent == null)
                parent = new EmptyParent();

            if(target.MissingAny("{", "}"))
                return Quick.Block(parent, target);

            var before = target.FirstToken("{");
            var after = target.LastToken("}");
            
            string codeInsideOutermostBraces;
            if (after.Length == 0)
            {
                var lengthToSlice = target.Length - before.Length - 2;
                if (lengthToSlice > 0)
                    codeInsideOutermostBraces = target.Substring(before.Length + 1, lengthToSlice);
                else
                    codeInsideOutermostBraces = target;
            }
            else
            {
                var targetLength = target.Length - before.Length - after.Length - 2;
                if (targetLength > 0)
                {
                    codeInsideOutermostBraces = target.Substring(before.Length + 1, targetLength);
                }
                else
                {
                    codeInsideOutermostBraces = "";
                }
            }
            
            var block = new Block(parent);
            if (noOuterBraces)
            {
                block.Before = null;
                block.After = null;
            }

            if (before.Trim().Length > 0)
                block.Children.AddLines(block, before.Lines().Trimmed().RemoveEmpty());

            if (codeInsideOutermostBraces.MissingAny("{", "}"))
            {
                // no further sub braces
                var statementBeforeOpeningBrace = block.Children[block.Children.Count - 1];
                block.Children.Remove(statementBeforeOpeningBrace);

                var block2 = Quick.Block(block, statementBeforeOpeningBrace.Line);
                block2.Children.AddLines(block, codeInsideOutermostBraces.Lines().Trimmed().RemoveEmpty());
                block.Children.Add(block2);

                if (after.Trim().Length > 0)
                    block.Children.AddLines(block, after.Lines().Trimmed().RemoveEmpty());
                return block;
            }

            var innerBlock = codeInsideOutermostBraces.AsBlock(block, false);
            block.Children.Add(innerBlock);
            return block;
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
                while (lineOfCode.ToLower().StartsWith(entry.X.ToLower()) &&
                       lineOfCode.ToLower().EndsWith(entry.Y.ToLower()))
                {
                    lineOfCode = lineOfCode
                        .Replace(entry.X, "")
                        .Replace(entry.Y, "");

                    lineOfCode = entry.Z + Environment.NewLine + lineOfCode + Environment.NewLine +
                                 entry.A;
                }

            return lineOfCode;
        }

        /// <summary>
        ///     When line contains X, replace with Y
        /// </summary>
        /// <param name="originalLineOfCode"></param>
        /// <returns></returns>
        public static string BlanketReplaceNow(string originalLineOfCode, string currentModuleName)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in BlanketReplacements)
            {
                var iterations = 0;
                while (++iterations < 100 && lineOfCode.ToLower().Contains(entry.X.ToLower()))
                {
                    lineOfCode = lineOfCode
                        .Replace(entry.X, entry.Y)
                        .Replace("CurrentModuleName", currentModuleName);
                }
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
            lineOfCode = lineOfCode.HandleTokenDefinitionOrder();

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

            var scrubbedLine = line.Replace("\r", "").Replace("\n", "").Trim();
            
            if (scrubbedLine.EndsWith(";")
                || scrubbedLine.EndsWith("{")
                || scrubbedLine.EndsWith("}"))
                return line;

            if (LineContainsAnyOfTheseThenShouldHaveASemiColon
                    .Any(x => scrubbedLine.Contains(x)) == false
                && LineDoesNotContainAnyOfThese.Any(x =>
                    scrubbedLine.Contains(x) == false 
                    && !scrubbedLine.EndsWith(":")))
            {
                if (nextLine.IsEmpty() || !nextLine.Contains("{"))
                    return line + ";";
            }
            if (scrubbedLine == ";")
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
                lines[i] = DetermineIfLineGetsASemicolon(lines[i],
                    i + 1 < lines.Count ? lines[i + 1] : null);

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
                while (lineOfCode.ToLower().EndsWith(entry.X.ToLower()))
                    lineOfCode = lineOfCode.Replace(entry.X.ToLower(), entry.Z);

            return lineOfCode;
        }

        public static string AllLinesNow(string translatedLine, Module sourceModule)
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
                result.AppendLine(Transform(line, sourceModule, nextLine));
            }

            var code = result.ToString();
            return code;
        }

        /// <summary>
        ///     Run all blanket and specialized replacements now
        /// </summary>
        /// <param name="originalLine"></param>
        /// <param name="currentModuleName"></param>
        /// <param name="nextLine"></param>
        /// <param name="translatedLine"></param>
        /// <returns></returns>
        public static string Transform(this string originalLine, Module sourceModule, string nextLine = null)
        {
            if (originalLine.IsEmpty())
                return "";

            if (originalLine.Trim() == "")
                return originalLine;

            if (originalLine.Trim().StartsWith("//"))
                return originalLine;

            var translatedLine = originalLine;
            translatedLine = translatedLine.HandleWith();

            translatedLine = BlanketReplaceNow(translatedLine, sourceModule.Name);
            translatedLine = StartsWithReplaceNow(translatedLine);
            translatedLine = EndsWithReplaceNow(translatedLine);

            translatedLine = StartsAndEndsWithReplaceNow(translatedLine);
            translatedLine = AboveAndBelowWithReplaceNow(translatedLine);

            translatedLine = WhenStartsWithReplaceOtherAndAppendNow(translatedLine);

            translatedLine = RegExReplacements.Shuffle(translatedLine);

            translatedLine = CleanupTranslatedLineOfCode(translatedLine);
            translatedLine = DetermineIfLineGetsASemicolon(translatedLine, nextLine);
            //translatedLine = BlanketReplaceNow(translatedLine, sourceModule.Name);
            
            if (translatedLine.Trim() == ";"
                || translatedLine.Trim() == ";;")
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
                    lineOfCode +=
                        " // TODO: Check: Too many iterations during StartsWithReplaceNow";
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

        /*
        public static bool FindCodeBetweenBraces(this string target, out string before,
            out string insideTheMostOuterBraces, out string after, out int indexOfMostInner)
        {
            var input = target;
            var regex = new Regex("{((?>[^{}]+|{(?<c>)|}(?<-c>))*(?(c)(?!)))"); //, RegexOptions.Singleline);
            var splits = regex.Split(input).Where(s => s.Length > 0).ToArray();

            if (splits.Length == 3)
            {
                before = splits[0] + "{";
                insideTheMostOuterBraces = splits[1];
                after = splits[2];
                indexOfMostInner = regex.Match(target, 0).Index;
                // indexOfMostInner = target.IndexOf(insideTheMostOuterBraces, StringComparison.InvariantCultureIgnoreCase);
                return true;
            }

            before = null;
            insideTheMostOuterBraces = target;
            after = null;
            indexOfMostInner = 0;
            return false;
        }
        */
    }
}