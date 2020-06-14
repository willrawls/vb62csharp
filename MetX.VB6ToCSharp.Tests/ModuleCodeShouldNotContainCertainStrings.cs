using System;
using System.Collections.Generic;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeShouldBeWellFormed
    {
        const string inputFilePath = @"I:\OneDrive\data\code\Slice and Dice\Sandy\CAssocItem.cls";
        const string outputPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC";

        static List<string> NoNos = new List<string>
        {
            "get;", "set;", "{;", "};", "';"
        };

        [TestMethod]
        public void In_CAssocItem_cs()
        {
            ICodeLine parent = new EmptyParent(-1); 
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(parent, inputFilePath, outputPath);
            
            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var noNoList = "\nShould not contain: \n";
            foreach (var noNo in NoNos)
            {
                if(code.Contains(noNo))
                    noNoList += $"  {noNo}\n";
            }

            Console.WriteLine("\n" + code);

            foreach (var noNo in NoNos)
            {
                Assert.IsFalse(code.Contains(noNo), noNoList);
            }
        }


        [TestMethod]
        public void WithBlockReplacementTest()
        {
            ICodeLine parent = new EmptyParent();


            var code =
@"With Fred
    .x = b
    .y = c
    .z = d
End With";
            var expected = 
@"    Fred.x;
    Fred.y = c;
    Fred.z = d
";
            var block = new Block(new EmptyParent(), "With Fred");
            block.Children.Add(new LineOfCode(block, "    .x = b"));
            block.Children.Add(new LineOfCode(block, "    .y = c"));
            block.Children.Add(new LineOfCode(block, "    .z = d"));
            block.Children.Add(new LineOfCode(block, "End With"));

            var actual = block.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }


    }
}