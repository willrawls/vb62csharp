using System;
using System.Linq;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Structure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class MassageTests
    {
        /*
        const string inputFilePath = @"I:\OneDrive\data\code\Slice and Dice\Sandy\CAssocItem.cls";
        const string outputPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC";

        static List<string> NoNos = new List<string>
        {
            "get;", "set;", "{;", "};", "';"
        };
        */

        [TestMethod]
        public void DetermineIfLineGetsASemicolonTest_PublicObjectFred_Brace()
        {
            var actual = Massage.DetermineIfLineGetsASemicolon("    public object Fred", "    {");

            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.Length > 0);

            Console.WriteLine(actual);
            Assert.IsFalse(actual.Contains("Fred;"));
            Assert.IsFalse(actual.Contains("{;"));
        }

        [DataTestMethod]
        [DataRow("012{4567}90", 0, 4, 3, 8, "012", "4567", "90")]
        [DataRow("012 { 4567 } 90", 0, 5, 4, 11, "012 ", " 4567 ", " 90")]
        [DataRow("1x { { y } } z", 0, 4, 3, 11, "1x ", " { y } ", " z")]

               //           1         2           3         4         5           6            7
               // 0123456789 123456789 1 2 3456789 123456789 123456789 123456 7 89 12345 6 78 9 012345
        [DataRow("2foreach(var y in z) {\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n}",
            // StartAt, Code, Open, Close
               0,       22,   21,   71, 
            "2foreach(var y in z) ", 
            "\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n", 
            "")]

               /*
               //           1         2           3         4         5           6            7      
               // 0123456789 123456789 1 2 3456789 123456789 123456789 123456 7 89 12345 6 78 9 012345
        [DataRow("3foreach(var y in z) {\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n}",
            50, 60, 59, 68,
            " if(a) \r\n",
            " b(); \r\n",
            "\r\n")]
        */
        public void FindMatching_Simple(string startingCode,
            int startLookingAtIndex, 
            int expectedIndexOfCode,
            int expectedIndexOfOpenBrace,
            int expectedIndexOfCloseBrace,
            
            string expectedBeforeOpenBrace,
            string expectedCodeFoundInsideBraces,
            string expectedAfterCloseBrace
            )
        {
            //var actual = CodeBetweenBraces.Factory(startingCode);
            var actual = startingCode.CodeBetweenBraces();

            Assert.IsNotNull(actual);
            
            var expected = new CodeBetweenBraces
            {
                StartingCode = startingCode,
                FindResult = true,
                IndexOfCode = expectedIndexOfCode,
                AfterCloseBrace = expectedAfterCloseBrace,
                CodeFoundInsideBraces = expectedCodeFoundInsideBraces,
                IndexOfOpenBrace = expectedIndexOfOpenBrace,
                IndexOfCloseBrace = expectedIndexOfCloseBrace,
                BeforeOpenBrace = expectedBeforeOpenBrace,
            };

            Assert.AreEqual(expected, actual, expected.Diff(actual));
        }

        [TestMethod]
        public void AsBlock_Simple()
        {
            var target =
                "foreach(var y in z) {\r\n    var someLine = ofCode;\r\n    if(a)\r\n{ b();\r\n}\r\n}";
            var expected = @"
    foreach(var y in z)
    {
        var someLine = ofCode;
        if(a)
        {
            b();
        }
    }
";
            var expectedLines = expected.Lines().Where(line => line.IsNotEmpty()).ToList();

            var parent = new EmptyParent();
            var block = target.AsBlock(parent, true);
            var actualCode = block.GenerateCode();
            var actualLines = actualCode.Lines().Where(line => line.IsNotEmpty()).ToList();

            Assert.AreEqual(expectedLines.Count, actualLines.Count);

            for (var index = 0; index < actualLines.Count; index++)
            {
                var actualLine = actualLines[index];
                var expectedLine = expectedLines[index];
                //var message = $"\r\nExpected:\r\n\t[[[{expectedLine}]]]\r\nActual:\r\n\t[[[{actualLine}]]]";
                Assert.AreEqual(expectedLine, actualLine);
            }
        }

        [TestMethod]
        public void Regexify_Simple()
        {
            var actual = "{}.[]()*a1".Regexify();
            Assert.AreEqual("\\{\\}\\.\\[\\]\\(\\)\\*a1", actual);
        }
    }
}