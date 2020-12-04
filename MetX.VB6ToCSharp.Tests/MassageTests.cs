using System;
using System.Collections.Generic;
using MetX.Scripts;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
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

        [DataRow("x { { y } } z", 0, " { y } ")]
        [DataRow("foreach(var y in z) {\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n}",  0, "\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n")]
        [DataRow("foreach(var y in z) {\r\n    var someLine = ofCode; if(a) \r\n{ b(); \r\n}\r\n}", 25, " b(); \r\n}")]
        [DataTestMethod]
        public void FindMatching_Simple(string target, int startAtIndex, string expected)
        {
            var actual = target.FindCodeBetweenBraces();
            Assert.AreEqual(expected, actual);
        }

        // code, before the most inner {, between { }, after the most inner
        [DataRow("x { { y } } z", "x {", " { y } ", "} z")]
        [DataTestMethod]
        public void FindCodeBetweenBraces_Simple(string target, string[] expected)
        {
            bool result = target.FindCodeBetweenBraces(out string before, out string mostInner, out string after);

            Assert.IsTrue(result);
            Assert.AreEqual(expected[0], before);
            Assert.AreEqual(expected[1], mostInner);
            Assert.AreEqual(expected[2], after);
        }

        [TestMethod]
        public void AsBlock_Simple()
        {
            var target = "foreach(var y in z) {\r\n    var someLine = ofCode;\r\n    if(a)\r\n{ b();\r\n}\r\n}";
            var parent = new EmptyParent();
            var expected = Quick.Block(parent, null);
            expected.Children.Add(Quick.Line(expected, "var someLine = ofCode;"));
            //expected.Children.Add(Quick.Line(expected, "if(a)"));
            var ifBlock = Quick.Block(expected, "if(a)");
            ifBlock.Children.Add(Quick.Line(ifBlock, "b();"));
            expected.Children.Add(ifBlock);

            var actualParent = new EmptyParent();
            var actual = target.FindCodeBetweenBraces().AsBlock(actualParent);

            Assert.IsTrue(expected.Equals(actual), 
                "\r\nExpected: " + expected.GenerateCode()
                + "Actual: " + actual.GenerateCode());
        }

        [TestMethod]
        public void Regexify_Simple()
        {
            string actual = "{}.[]()*a1".Regexify();
            Assert.AreEqual("\\{\\}\\.\\[\\]\\(\\)\\*a1", actual);
        }
    }
}