using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeBlockTests : AbstractCodeBlock

    {
        [TestMethod]
        public void ProperIndentation3LevelsDeep()
        {
            var codeBlock = new CodeBlock(this, "Fred");

            codeBlock.Children.Add(new CodeBlock(codeBlock, "One"));

            AbstractCodeBlock two = new CodeBlock(codeBlock, "Two");
            codeBlock.Children.Add(two);
            codeBlock.Children.Add(new CodeBlock(codeBlock, "Three"));
            two.Children.Add(new CodeBlock(two, "Four"));

            AbstractCodeBlock five = new CodeBlock(two, "Five");
            two.Children.Add(five);

            five.Children.Add(new CodeBlock(five, "Six"));
            five.Children.Add(new CodeBlock(five, "Seven"));

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