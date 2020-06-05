using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CSharpPropertyTests
    {
        [TestMethod]
        public void SimpleTest()
        {
            ICodeLine parent = new EmptyCodeParent();

            var target = new CSharpProperty(parent, 1)
            {
                Comment = "' TheComment",
                Name = "F",
                Type = "G",
            };

            target.Get.Line = "A";
            target.Get.Encountered = true;
            target.Get.Children = new List<ICodeLine>
            {
                new CodeBlock(target.Get, "B"),
                new CodeBlock(target.Get, "C")
            };

            target.Set.Line = "D";
            target.Set.Encountered = true;
            target.Set.Children = new List<ICodeLine>();
            target.Set.Children.Add(new CodeBlock(target.Get, "E"));

            var expected =
@"    // TheComment
    public G F
    {
        get
        {
            A
            {
                B
                C
            }
        }
        set
        {
            D
            {
                E
            }
        }
";
            var actual = target.GenerateCode();
            Console.WriteLine(actual);

            Assert.AreEqual(expected, actual);
        }
    }
}