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
            var block = _.B(_.Top(0), "A", "B");
            Assert.AreEqual("    ", block.Indentation);
            Assert.AreEqual("        ", block.SecondIndentation);
        }

        
        [TestMethod]
        public void ProperIndentation3LevelsDeep()
        {
            ICodeLine parent = new EmptyParent(0);

            var block1 = _.B(parent, "Fred", "One");

            AbstractBlock block2 = _.B(block1, "Two");
            block1.Children.Add(block2);
            block1.Children.Add(_.B(block1, "Three"));
            block2.Children.Add(_.B(block2, "Four"));

            AbstractBlock block5 = _.B(block2, "Five");
            block2.Children.Add(block5);

            block5.Children.Add(_.B(block5, "Six"));
            block5.Children.Add(_.B(block5, "Seven", "Eight"));

            const string expected =
@"    Fred
    {
        One;
        Two
        {
            Four
            {
            }
            Five
            {
                Six
                {
                }
                Seven
                {
                    Eight;
                }
            }
        }
        Three
        {
        }
    }
";

            var actual = block1.GenerateCode();

            Console.WriteLine(actual);
            Assert.AreEqual("\n" + expected, "\n" + actual);
        }

        [TestMethod]
        public void BlockWithLineAndSubLineTest()
        {
            ICodeLine parent = new EmptyParent();
            var block = _.B(parent, "Fred", "George");

            var expected = "    Fred\r\n    {\r\n        George;\r\n    }\r\n";
            var actual = block.GenerateCode();
            Assert.AreEqual("\n" + expected, "\n" + actual);
        }
    }
}