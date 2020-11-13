using System;
using System.Collections.Generic;
using System.Diagnostics;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CAssocItem_CodeShouldBeWellFormed
    {
        const string InputFilePath = @"CAssocItem.cls";

        [TestMethod]
        public void ProperlyTranslate_CAssocItem_cs()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(InputFilePath);

            // Must have
            var message = LookFor(code, true, new []
            {
                "            return sGetToken",
                "            return m_sValue;", 
                "m_", 
                "            public string Key;",
                // "        {\n}",
                "public string Value",
            });

            // Must not have
            message += LookFor(code, false, new []
            {
                "// freeware",
                "get;", "set;", 
                "{;", "};", 
                "';", "*;", 
                "F = ", 
                // "Value = ", 
                "assumed;", 
                "risk.         //",
                "        ;",
                "delimiter.         //",
            });

            message += CheckOccurrences(code, new[]
            {
                "return m_sValue;",
            });

            Assert.IsTrue(message == string.Empty,
                $"\n==========\n{message}\n==========\n{code}");

            Console.WriteLine(code);
        }

        public static string CheckOccurrences(string code, string[] list, int minimum = 1, int maximum = -1)
        {
            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var i = 0;
            var message = "";
            if (maximum < minimum)
                maximum = minimum;

            foreach (var item in list)
            {
                var tokenCount = code.TokenCount(item);
                if (tokenCount < minimum)
                    message += $"{++i}:  To few of \"{item}\"\n";
                if (tokenCount > maximum)
                    message += $"{++i}:  To many of \"{item}\"\n";
            }

            return message == "" 
                ? "" 
                : $"----------\nOccurrences other than expected:\n{message}\n";
        }

        public static string LookFor(string code, bool mustHave, string[] list)
        {
            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var i = 0;
            var message = "";
            if(mustHave)
            {
                foreach (var item in list)
                    if (!code.Contains(item))
                        message += $"{++i}:  {item}\n";
            }
            else
            {
                foreach (var item in list)
                    if (code.Contains(item))
                        message += $"{++i}:  {item}\n";
            }

            return message == "" 
                ? "" 
                : $"----------\nCode {(mustHave ? "is missing" : "must not have")}:\n{message}\n";
        }


        /*[TestMethod, Ignore("Never going to happen this way")]
        public void WithBlockReplacementTest()
        {
            var code =
@"With Fred
    .x = b
    .y = c
    .z = d
End With";
            var expected = 
@"    Fred.x;
    Fred.y = c;
    Fred.z = d;
";
            var parent = _.Top();
            var block = new Block(parent, "With Fred");
            block.Children.Add(new LineOfCode(block, "    .x = b"));
            block.Children.Add(new LineOfCode(block, "    .y = c"));
            block.Children.Add(new LineOfCode(block, "    .z = d"));
            block.Children.Add(new LineOfCode(block, "End With"));

            var actual = block.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }*/


    }
}