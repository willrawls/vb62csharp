using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public static class RegExReplacements
    {
        // something(x,y) => x.somethingElse(y)
        public static List<XReplace> toX_function_Y = new List<XReplace>
        {
            new XReplace("instr", "$1.Contains($2)"),
            new XReplace( "left", "$1.Substring(0, $2)"),
            new XReplace(@"left\$", "$1.Substring(0, $2)"),
            new XReplace("right", "$1.Substring($1.Length - $2)"),
            new XReplace(@"right\$", "$1.Substring($1.Length - $2)"),
        };

        // something(x) => x.OtherSomething()
        public static List<XReplace> toX_function = new List<XReplace>
        {
            new XReplace("ucase", "$1.ToUpper()"),
            new XReplace(@"ucase\$", "$1.ToUpper()"), // Since it's going into a regular expression, the entry must meet its rules
        };

        // something x  ==> something(x)
        public static List<string> fromBare_to_function_X_ = new List<string>
        {
            "Add",
            "AddItem",
            "SetListIndex",
        };

        static RegExReplacements()
        {
            // something(x,y) => x.somethingElse(y)
            foreach (var wrappable in toX_function_Y)
            {
                // When "a(x,y) => x.A(y)   (for instance Z = Contains)
                Library.Add(new RegexReplace(
                    new Regex($@".*{wrappable.X}\s*\(\s*(?<X>.*)\s*,\s*(?<Y>.*)\b\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                    wrappable.Y));
            }

            // something x  ==> something(x)
            foreach (var wrappable in fromBare_to_function_X_)
            {
                Library.Add(new RegexReplace(
                    new Regex(wrappable + @" (.+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                    wrappable + "($1)"));
            }

            // something(x) => x.OtherSomething()
            foreach (var wrappable in toX_function)
            {
                var item = new RegexReplace(
                    new Regex(wrappable.X + @"\((.+)\)",
                        RegexOptions.IgnoreCase | RegexOptions.Compiled),
                    wrappable.Y);
                Library.Add(item);
            }
        }

        public static List<RegexReplace> Library = new List<RegexReplace>
        {
            // X = X + ...
            new RegexReplace(
                new Regex(@"(.*) = (\1 [+-\\])(.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1 += $3;"),

            // For x = y To z
            new RegexReplace(
                new Regex(@"For (.+) = (.*) To (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "for(var $1 = $2; $1 < $3; $1++) //SOB//"),

            // For Each x In y
            new RegexReplace(
                new Regex(@"For Each (.+) In (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "foreach(var $1 in $2) //SOB//"),

            // Mid$(x,y) ---
            new RegexReplace(
                new Regex(@"Mid\$*\((.+),(.+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1.Substring($2)"),

            // ... x As y
            new RegexReplace(
                new Regex(@"(.+) As (.+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$2 $1;"),

            // ... x As y
            new RegexReplace(
                new Regex(@"(.+) As (.+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$2 $1;"),

            // Do While x > y ---
            new RegexReplace(
                new Regex(@"Do While (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "while($1)//SOB//"),

            // Do Until x > y ---
            new RegexReplace(
                new Regex(@"Do Until (.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "while($1)//SOB// // 'Until'"),
            
            // UBound(x) ---
            new RegexReplace(
                new Regex(@"UBound\$*\((.+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "$1.Length"),
            
            // LBound(x) ---
            new RegexReplace(
                new Regex(@"LBound\$*\((.+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
                "0 /* $1.Length */"),
        };

        public static string Shuffle(string line)
        {
            foreach (var entry in Library)
            {
                line = entry.Regex.Replace(line, entry.ReplacePattern);
            }

            return line;
        }
    }
}