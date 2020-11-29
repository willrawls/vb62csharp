using System;
using System.Collections.Generic;
using System.Text;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeLineRelatedTests
    {
        [TestMethod]
        public void Indentifier_Simple()
        {
            var top = Quick.Top();
            var block = Quick.Block(top, "A", "B");
            
            block.ResetIndent(0);
            Assert.AreEqual(1, block.Children.Count, block.GenerateCode());

            block.Children.Add(Quick.Line(block, "Fred"));
            Assert.AreEqual(1, block.Children[0].Indent);
        }

        
        [TestMethod]
        public void ProperIndentation3LevelsDeep()
        {
            ICodeLine parent = new EmptyParent(0);

            var block1 = Quick.Block(parent, "Fred", "One;");

            AbstractBlock block2 = Quick.Block(block1, "Two");
            block1.Children.Add(block2);
            block1.Children.Add(Quick.Block(block1, "Three"));
            block2.Children.Add(Quick.Block(block2, "Four"));

            AbstractBlock block5 = Quick.Block(block2, "Five");
            block2.Children.Add(block5);

            block5.Children.Add(Quick.Block(block5, "Six"));
            block5.Children.Add(Quick.Block(block5, "Seven", "Eight;"));

            const string expected =
@"    Fred
    {
        One;
        Two
        {
            Four
            Five
            {
                Six
                Seven
                {
                    Eight;
                }
            }
        }
        Three
    }
";

            block1.ResetIndent(1);
            var actual = block1.GenerateCode();

            Console.WriteLine(actual);
            Assert.AreEqual("\n" + expected, "\n" + actual);
        }

        [TestMethod]
        public void BlockWithLineAndSubLineTest()
        {
            ICodeLine parent = new EmptyParent();
            var block = Quick.Block(parent, "Fred", "George;");

            var expected = "    Fred\r\n    {\r\n        George;\r\n    }\r\n";
            
            block.ResetIndent(1);
            var actual = block.GenerateCode();
            
            Assert.AreEqual("\n" + expected, "\n" + actual);
        }
    }
}