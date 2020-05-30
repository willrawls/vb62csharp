using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public static class Massage
    {
        /// <summary>
        /// When line starts with X and ends with Y:
        ///      Remove X and Y
        ///      Add Z as a line above the beginning
        ///      Add A as a line below the end
        /// </summary>
        public static List<Alice> AboveAndBelowReplacements { get; set; } = new List<Alice>
        {
        };

        /// <summary>
        /// When line contains X:
        ///      Replace with Y
        /// </summary>
        public static List<Alice> BlanketReplacements { get; set; } = new List<Alice>
        {
            new Alice("Exit Property", "return; // ???"),
            new Alice("Exit Function", "return; // ???"),
            new Alice("Exit Sub", "return;"),
            new Alice("For Each ", "foreach( var "),
            new Alice(" In ", " in "),
            new Alice(";;", ";"),
            new Alice("; = ", " = "),
            new Alice("{ get; set; }", ";"),
            new Alice("Static", "static"),
            new Alice("String", "string"),
            new Alice("Collection", "Dictionary<string,string>()"),
            new Alice("True", "true"),
            new Alice("False", "false"),
            new Alice("private ", "public "), // Because I believe everything should be public
            new Alice("Err.Clear", string.Empty),
            new Alice("Err.Number", "ex"),
            new Alice("Me.", "this."),
        };

        /// <summary>
        /// When line ends with X:
        ///      Replace X with Y
        /// </summary>
        public static List<Alice> EndsWithReplacements { get; set; } = new List<Alice>
        {
            new Alice(";;", ";"),
        };

        /// <summary>
        /// When line starts with X and ends with Y:
        ///      Replace Z with A
        /// </summary>
        public static List<Alice> StartsAndEndsWithReplacements { get; set; } = new List<Alice>()
        {
            new Alice("Set", "Nothing", "Set", ""),

        };

        /// <summary>
        /// When line starts with X, Replace X with Y, Replace Z with A
        /// </summary>
        public static List<Alice> StartsWithReplacements { get; set; } = new List<Alice>
        {
            new Alice("' ", "// "),
            new Alice("On Error GoTo ", "// TODO: Rewrite try/catch and/or goto. "),
        };

        /// <summary>
        /// When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static List<Alice> WhenStartsWithReplaceOtherReplacements { get; set; } = new List<Alice>()
        {
            new Alice("foreach(", null, null, " )"),
        };

        /// <summary>
        /// When line starts with X and ends with Y:
        ///      Remove X and Y, 
        ///      Add Z as a line above the beginning, 
        ///      Add A as a line below the end
        /// </summary>
        public static string AboveAndBelowWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode;
            foreach (var entry in AboveAndBelowReplacements)
            {
                while (lineOfCode.ToLower().StartsWith(entry.X.ToLower()) && lineOfCode.ToLower().EndsWith(entry.Y.ToLower()))
                {
                    lineOfCode = lineOfCode
                        .Replace(entry.X, "")
                        .Replace(entry.Y, "");

                    lineOfCode = entry.Z + Environment.NewLine + lineOfCode + Environment.NewLine + entry.A;
                }
            }

            return lineOfCode;
        }

        /// <summary>
        /// When line contains X, replace with Y
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
        /// Special case replacements
        /// </summary>
        /// <param name="lineOfCode"></param>
        /// <returns></returns>
        public static string CleanupTranslatedLineOfCode(string lineOfCode)
        {
            if (lineOfCode.Contains("foreach(") && !lineOfCode.Contains(")"))
                lineOfCode += " )";

            if (lineOfCode.Contains("foreach(") && lineOfCode.EndsWith(");"))
                lineOfCode = lineOfCode.FirstToken(");") + ")";

            if (lineOfCode.StartsWith("Next "))
                lineOfCode = "}\r\n";

            return lineOfCode.Trim();
        }

        /// <summary>
        /// Adds semicolon to end of appropriate lines
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static IList<string> DetermineWhichLinesGetASemicolon(IList<string> lines)
        {
            if (lines.IsEmpty())
                return lines;

            if (lines.Count == 1)
            {
                lines[0] += ";";
                return lines;
            }

            for (var i = 0; i < lines.Count; i++)
            {
                if (lines[i].IsNotEmpty()
                    && !lines[i].EndsWith(";")
                    && !lines[i].EndsWith("{")
                    && !lines[i].EndsWith("}")
                )
                    lines[i] += ";";
            }

            return lines;
        }

        /// <summary>
        /// When line ends with X
        ///      Replace X with Z
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

            StringBuilder result = new StringBuilder();
            List<string> lines = translatedLine
                .Replace("\r", string.Empty)
                .Split('\n').ToList();

            foreach (var line in lines) 
                result.AppendLine(Now(line));
            var code = result.ToString();
            return code;
        }
        /// <summary>
        /// Run all blanket and specialized replacements now
        /// </summary>
        /// <param name="translatedLine"></param>
        /// <returns></returns>
        public static string Now(string translatedLine)
        {
            translatedLine = BlanketReplaceNow(translatedLine);
            translatedLine = StartsWithReplaceNow(translatedLine);
            translatedLine = EndsWithReplaceNow(translatedLine);

            translatedLine = StartsAndEndsWithReplaceNow(translatedLine);
            translatedLine = AboveAndBelowWithReplaceNow(translatedLine);

            translatedLine = WhenStartsWithReplaceOtherAndAppendNow(translatedLine);

            translatedLine = CleanupTranslatedLineOfCode(translatedLine);

            return translatedLine;
        }

        /// <summary>
        /// When line starts with X and ends with Y
        ///      Replace Z with A
        /// </summary>
        public static string StartsAndEndsWithReplaceNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode.Trim();
            foreach (var entry in StartsAndEndsWithReplacements)
                if (lineOfCode.ToLower().StartsWith(entry.X.ToLower()) && lineOfCode.ToLower().EndsWith(entry.Y.ToLower()))
                    lineOfCode = lineOfCode.Replace(entry.Z, entry.A);

            return lineOfCode;
        }
        
        /// <summary>
        /// When line starts with X, Replace X with Y, Replace Z with A
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
                    if(entry.Y.IsNotEmpty())
                        lineOfCode = lineOfCode.Replace(entry.X, entry.Y);
                    if(entry.Z.IsNotEmpty() && entry.A.IsNotEmpty())
                        lineOfCode = lineOfCode.Replace(entry.Z, entry.A);
                }
                if(iterations > 98)
                    throw new Exception("Too many tries in StartsWithReplaceNow");
            }

            return lineOfCode;
        }

        /// <summary>
        /// When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static string WhenStartsWithReplaceOtherAndAppendNow(string originalLineOfCode)
        {
            if (originalLineOfCode.IsEmpty())
                return originalLineOfCode;

            var lineOfCode = originalLineOfCode.Trim();
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
    }
}