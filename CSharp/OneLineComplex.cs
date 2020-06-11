using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class OneLineComplex
    {
        public static List<RegexReplace> Shuffle = new List<RegexReplace>
        {
            // Instr
            new RegexReplace(
                new Regex(@".*instr\s*\(\s*(?<X>.*)\s*,\s*(?<Y>.*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1.Contains($2"),

            // X = X + ...
            new RegexReplace(
                new Regex(@"(.*) = (\1 [+-\\])(.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1 += $3;"),

            // For x = y To z
            new RegexReplace(
                new Regex(@"For (.+) = (.*) To (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "for(var $1 = $2; $1 < $3; $1++) //SOB//"),

            // Add x
            new RegexReplace(
                new Regex(@"Add (.+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "Add($1)"),

            // Mid$(x,y) ---
            new RegexReplace(
                new Regex(@"Mid\$*\((.+),(.+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1.Substring($2)"),

            // Do While x > y ---
            new RegexReplace(
                new Regex(@"Do While (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "while($1)//SOB//"),
            
        };

        public static string Now(string line)
        {
            foreach (var entry in Shuffle) 
                line = entry.Regex.Replace(line, entry.ReplacePattern);

            return line;
        }
    }
}