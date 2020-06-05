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
        public void SimpleGetSetTest()
        {
            ICodeLine parent = new EmptyParent();

            // ReSharper disable once UseObjectOrCollectionInitializer
            var target = new CSharpProperty(parent)
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

        [TestMethod]
        public void SimplePropertyPartTest()
        {
            ICodeLine parent = new EmptyParent();
            var part = new CSharpPropertyPart(parent, PropertyPartType.Get)
            {
                Line = "Testing",
                PartType = PropertyPartType.Get,
                Encountered = true
            };
            part.Children = new List<ICodeLine>
            {
                new CodeLine(part, "123")
            };

            var actual = part.GenerateCode();
            Assert.AreEqual("\n    get\r\n    {\r\n        123\r\n    }\r\n", "\n" + actual);
        }
    }
}