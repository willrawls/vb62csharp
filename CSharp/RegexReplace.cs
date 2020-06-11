using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public class RegexReplace
    {
        public Regex Regex;
        public string ReplacePattern;

        public RegexReplace(Regex regex, string replacePattern)
        {
            Regex = regex;
            ReplacePattern = replacePattern;
        }

        public string Replace(string line) => Regex.Replace(line, ReplacePattern);
    }
}