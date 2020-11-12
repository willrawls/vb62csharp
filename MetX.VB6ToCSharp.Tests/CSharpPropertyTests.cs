using System;
using System.Collections.Generic;
using System.Reflection;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CSharpPropertyTests
    {
        [TestMethod]
        public void NoChildren_ToEmptyGetAndSet()
        {
            var target = QuickCSharpProperty("F", "G");

            const string expected = "    public G F { get; set; }\r\n";
            var actual = target.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }

        public static CSharpProperty QuickCSharpProperty(string name = null, string type = null, string line = null)
        {
            ICodeLine parent = new EmptyParent();
            var target = new CSharpProperty(parent)
            {
                Name = name,
                Type = type,
                Line = line,
            };
            return target;
        }

        [TestMethod]
        public void NoChildrenWithLine()
        {
            var target = QuickCSharpProperty("F", "G", "H");
            
            const string expected = "    H\r\n    public G F { get; set; }\r\n";
            var actual = target.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }


        [TestMethod]
        public void Comment_GetWith2Children_SetWith1Child()
        {
            ICodeLine parent = new EmptyParent();

            // ReSharper disable once UseObjectOrCollectionInitializer
            var target = new CSharpProperty(parent)
            {
                Comment = "' TheComment",
                Name = "F",
                Type = "G",
            };

            target.Get.Encountered = true;
            target.Get.Line = "A";
            target.Get.Children.AddRange(
                new List<ICodeLine>()
                {
                    _.Line(target.Get, "B"),
                    _.Line(target.Get, "C")
                });

            target.Set.Line = "D";
            target.Set.Encountered = true;
            target.Set.Children.Add(_.Line(target.Set, "E"));

            var expected =
@"    //  TheComment
    public G F
    {
        get
        {
            A
            {
                B;
                C;
            }
        }

        set
        {
            D
            {
                E;
            }
        }

    }
";
            var actual = target.GenerateCode();
            Extensions.AreEqualFormatted(expected, actual);
        }

        [TestMethod]
        public void WithGetWithTwoLines()
        {
            ICodeLine parent = new EmptyParent();
            var get = new CSharpPropertyPart(parent, PropertyPartType.Get)
            {
                PartType = PropertyPartType.Get,
                Encountered = true,
            };
            get.Children.Add(_.Line(get, "Testing"));
            get.Children.Add(_.Line(get, "123"));

            var actual = get.GenerateCode();
            var expected = "    get\r\n    {\r\n        Testing;\r\n        123;\r\n    }\r\n";
            Extensions.AreEqualFormatted(expected, actual);
        }
    }
}


