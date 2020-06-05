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
            AbstractBlock parent = new EmptyParent();

            var codeBlock = new Block(parent, "Fred");

            codeBlock.Children.Add(new Block(codeBlock, "One"));

            AbstractBlock two = new Block(codeBlock, "Two");
            codeBlock.Children.Add(two);
            codeBlock.Children.Add(new Block(codeBlock, "Three"));
            two.Children.Add(new Block(two, "Four"));

            AbstractBlock five = new Block(two, "Five");
            two.Children.Add(five);

            five.Children.Add(new Block(five, "Six"));
            five.Children.Add(new Block(five, "Seven"));

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