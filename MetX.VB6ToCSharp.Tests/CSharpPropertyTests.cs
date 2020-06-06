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

            //target.Get.Line = "A";
            target.Get.Encountered = true;
            target.Get.Blocks.AddRange(
                new List<ICodeLine>()
                {
                    Block.New(target.Get, "B"),
                    Block.New(target.Get, "C")
                });

            target.Set.Line = "D";
            target.Set.Encountered = true;
            target.Set.Blocks.Add(new Block(target.Set, "E"));

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
            var get = new CSharpPropertyPart(parent, PropertyPartType.Get)
            {
                Line = "Testing",
                PartType = PropertyPartType.Get,
                Encountered = true,
                Blocks = new List<ICodeLine>
                {
                    new LineOfCode(parent, "123")
                }
            };

            var actual = get.GenerateCode();
            Assert.AreEqual("\n    get\r\n    {\r\n        Testing\r\n        123\r\n    }\r\n", "\n" + actual);
        }
    }
}