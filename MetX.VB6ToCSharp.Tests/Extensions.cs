using System;
using MetX.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    public static class Extensions
    {
        public static void AreEqualFormatted(object expected, object actual, string message = null)
        {
            expected = "\n" + expected;
            actual = "\n" + actual;
            Assert.AreEqual(expected, actual, message);
            Console.WriteLine("Expected:" + expected);
            Console.WriteLine("Actual  :" +  actual);
        }

        public static string GetRidOfEmptyLines(this string target)
        {
            if (target.IsEmpty())
                return "";

            var output = "";
            foreach (var line in target.Lines())
            {
                if (line.Trim().Length != 0)
                {
                    output += line + "\r\n";
                }
            }

            if (output.IsNotEmpty())
                return "\r\n" + output;

            return output;
        }
    }
}