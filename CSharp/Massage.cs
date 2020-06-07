using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class Massage
    {
        /// <summary>
        /// When line starts with X and ends with Y:
        ///      Remove X and Y
        ///      Add Z as a line above the beginning
        ///      Add A as a line below the end
        /// </summary>
        public static List<FourWayReplace> AboveAndBelowReplacements { get; set; } = new List<FourWayReplace>
        {
        };

        /// <summary>
        /// When line contains X:
        ///      Replace with Y
        /// </summary>
        public static List<FourWayReplace> BlanketReplacements { get; set; } = new List<FourWayReplace>
        {
            new FourWayReplace("Exit Property", "return; // ???"),
            new FourWayReplace("Exit Function", "return; // ???"),
            new FourWayReplace("Exit Sub", "return;"),
            new FourWayReplace("For Each ", "foreach( var "),
            new FourWayReplace(" In ", " in "),
            new FourWayReplace(";;", ";"),
            new FourWayReplace("; = ", " = "),
            new FourWayReplace("{ get; set; }", ";"),
            new FourWayReplace("Static", "static"),
            new FourWayReplace("String", "string"),
            new FourWayReplace("Collection", "Dictionary<string,string>()"),
            new FourWayReplace("using System.Dictionary<string,string>()s", "using System.Collections"),
            new FourWayReplace("True", "true"),
            new FourWayReplace("False", "false"),
            new FourWayReplace("private ", "public "), // Because I believe everything should be public
            new FourWayReplace("Err.Clear", string.Empty),
            new FourWayReplace("Err.Number", "ex"),
            new FourWayReplace("Me.", "this."),
            new FourWayReplace("/*;", "/*"),
            new FourWayReplace("*/;", "*/"),
            new FourWayReplace("); ); )", ")"),
            new FourWayReplace("); )", ")"),
        };

        /// <summary>
        /// When line ends with X:
        ///      Replace X with Y
        /// </summary>
        public static List<FourWayReplace> EndsWithReplacements { get; set; } = new List<FourWayReplace>
        {
            new FourWayReplace(";;", ";"),
            new FourWayReplace(":;", "~~~"),
        };

        /// <summary>
        /// When line starts with X and ends with Y:
        ///      Replace Z with A
        /// </summary>
        public static List<FourWayReplace> StartsAndEndsWithReplacements { get; set; } = new List<FourWayReplace>()
        {
            new FourWayReplace("Set", "Nothing", "Set", ""),

        };

        /// <summary>
        /// When line starts with X, Replace X with Y, Replace Z with A
        /// </summary>
        public static List<FourWayReplace> StartsWithReplacements { get; set; } = new List<FourWayReplace>
        {
            new FourWayReplace("' ", "// "),
            new FourWayReplace("On Error GoTo ", "// TODO: Rewrite try/catch and/or goto. "),
        };

        /// <summary>
        /// When a line starts with X, Replace Y with Z, Append A
        /// </summary>
        public static List<FourWayReplace> WhenStartsWithReplaceOtherReplacements { get; set; } = new List<FourWayReplace>()
        {
            new FourWayReplace("foreach(", null, null, " )"),
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
        public static string CleanupTranslatedLineOfCode(string originalLineOfCode)
        {
            var lineOfCode = originalLineOfCode;
            if (lineOfCode.TokenCount() == 4 && lineOfCode.TokenAt(3) == "As")
            {
                lineOfCode = $"{lineOfCode.TokenAt(4)} {lineOfCode.TokenAt(2)};";
            }

            if (lineOfCode.Contains("foreach(") && !lineOfCode.Contains(")"))
                lineOfCode += " )";

            if (lineOfCode.Contains("foreach(") && lineOfCode.EndsWith(");"))
                lineOfCode = lineOfCode.FirstToken(");") + ")";

            if (lineOfCode.StartsWith("Next "))
                lineOfCode = "}\r\n";

            if (lineOfCode == "'" || lineOfCode == ";" || lineOfCode == "End With")
                lineOfCode = "";

            lineOfCode = lineOfCode
                .Replace("\r", "")
                .Replace("\n\n", "\n")
                .Replace("\n\n", "\n")
                .Replace("\n\n", "\n");

            return lineOfCode.Trim();
        }

        /// <summary>
        /// Adds semicolon to end of appropriate lines
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static string DetermineIfLineGetsASemicolon(string line)
        {
            if (line.IsEmpty() || line.EndsWith(";"))
                return line;

            if (!line.Trim().EndsWith(";") && !line.Trim().EndsWith("{") && !line.Trim().EndsWith("}"))
                line += ";";

            return line;
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

            var result = new StringBuilder();
            var lines = translatedLine
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
            translatedLine = DetermineIfLineGetsASemicolon(translatedLine);

            return translatedLine.Trim();
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