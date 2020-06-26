using System;
using System.Collections.Generic;
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
        const string inputFilePath = @"CAssocItem.cls";
        const string outputPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC";

        [TestMethod]
        public void ProperlyTranslate_CAssocItem_cs()
        {
            List<string> musts = new List<string>
            {
                "// freeware", 
                "            return sGetToken",
                "            return m_sValue;", 
                "m_", 
                "            public string Key;",
                "        {\n}",
                "public string Value",
            };

            List<string> noNos = new List<string>
            {
                "get;", "set;", 
                "{;", "};", 
                "';", "*;", 
                "F = ", 
                "Value = ", 
                "assumed;", 
                "risk.         //",
                "        ;",
            };

            ICodeLine parent = new EmptyParent(-1); 
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(parent, inputFilePath, outputPath);
            
            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var noNoList = "\nShould not contain:\n";
            foreach (var noNo in noNos)
            {
                if(code.Contains(noNo))
                    noNoList += $"  {noNo}\n";
            }

            var mustList = "\nMust contain:\n";
            foreach (var must in musts)
            {
                if(!code.Contains(must))
                    mustList += $"  {must}\n";
            }

            mustList += "\r\n";
            Console.WriteLine("\n" + code);

            foreach (var noNo in noNos) 
                Assert.IsFalse(code.Contains(noNo), noNoList + "\n" + mustList);

            foreach (var must in musts) 
                Assert.IsTrue(code.Contains(must), mustList);
        }


        [TestMethod, Ignore("Never going to happen this way")]
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
        }


    }
}