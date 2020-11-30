using System.Linq;
using MetX.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    public class CheckConvertedCode
    {
        public static string CheckOccurrences(string code, int minimum, string[] list,
            int maximum = -1)
        {
            if (list.IsEmpty())
                return "";

            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var i = 0;
            var message = "";
            if (maximum < minimum)
                maximum = minimum;

            foreach (var item in list.Where(x => x.IsNotEmpty()))
            {
                var tokenCount = code.TokenCount(item);
                var itemCount = tokenCount - 1;
                if (itemCount < minimum)
                    message += $"{++i}:  To few of \"{item}\"\n";
                if (itemCount > maximum)
                    message += $"{++i}:  To many of \"{item}\"\n";
            }

            return message == ""
                ? ""
                : $"----------\nOccurrences other than expected:\n{message}\n";
        }

        public static string LookFor(string code, bool mustHave, string[] list)
        {
            if (list.IsEmpty())
                return "";

            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var i = 0;
            var message = "";
            if (mustHave)
            {
                foreach (var item in list.Where(x => x.IsNotEmpty()))
                    if (!code.Contains(item))
                        message += $"{++i}:  {item}\n";
            }
            else
            {
                foreach (var item in list.Where(x => x.IsNotEmpty()))
                    if (code.Contains(item))
                        message += $"{++i}:  {item}\n";
            }

            return message == ""
                ? ""
                : $"----------\nCode {(mustHave ? "is missing" : "must not have")}:\n{message}\n";
        }
    }
}