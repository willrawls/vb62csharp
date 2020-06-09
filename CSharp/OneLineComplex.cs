using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class OneLineComplex
    {
        //public static string InstrPattern = @".*instr\s*\(\s*(?<X>.*)\s*,\s*(?<Y>.*)\b";
        //public static Regex Instr = new Regex(InstrPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static List<RegexReplace> Shuffle = new List<RegexReplace>
        {
            new RegexReplace(
                new Regex(@".*instr\s*\(\s*(?<X>.*)\s*,\s*(?<Y>.*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1.Contains($2"),

        };


        public static string ReplaceWithRegex(this RegexReplace find, string toSearch)
        {
            var answer = find.Regex.Replace(toSearch, "x.Contains($3,$2)");
            return answer;
        }
    }
}