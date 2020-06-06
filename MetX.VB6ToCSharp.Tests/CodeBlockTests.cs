using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeBlockTests
    {
        [TestMethod]
        public void ProperIndentation3LevelsDeep()
        {
            ICodeLine parent = new EmptyParent();

            var codeBlock = new Block(parent, "Fred");

            codeBlock.Blocks.Add(new Block(codeBlock, "One"));

            AbstractBlock two = new Block(codeBlock, "Two");
            codeBlock.Blocks.Add(two);
            codeBlock.Blocks.Add(new Block(codeBlock, "Three"));
            two.Blocks.Add(new Block(two, "Four"));

            AbstractBlock five = new Block(two, "Five");
            two.Blocks.Add(five);

            five.Blocks.Add(new Block(five, "Six"));
            five.Blocks.Add(new Block(five, "Seven"));

            const string expected =
@"    Fred
    {
        One
        Two
        {
            Four
            Five
            {
                Six
                Seven
            }
        }
        Three
    }
";

            var actual = codeBlock.GenerateCode();

            Console.WriteLine(actual);
            Assert.AreEqual(expected, actual);
        }
    }
}