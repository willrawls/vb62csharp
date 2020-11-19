using System.Collections.Generic;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CSharpPropertyPartTests
    {
        [TestMethod]
        public void GetWithLine()
        {
            var parent = new EmptyParent();
            var target = QuickCSharpPropertyPart(parent, "A");

            const string expected = "    get\r\n    {\r\n        A\r\n        {\r\n        }\r\n    }\r\n";
            var actual = target.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }

        [TestMethod]
        public void GetWithLineAnd1Child()
        {
            var parent = new EmptyParent();
            CSharpPropertyPart target = QuickCSharpPropertyPart(parent, "var x = new A");
            target.Children.Add(Quick.Line(target, "C = 'c',"));

            const string expected = "    get\r\n    {\r\n        var x = new A\r\n        {\r\n            C = 'c',\r\n        }\r\n    }\r\n";
            var actual = target.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }

        [TestMethod]
        public void GetWithBlock()
        {
            var parent = CSharpPropertyTests.QuickCSharpProperty();
            
            parent.Get.Children.Add(Quick.Line(parent.Get, "string sAllKeyValues;"));

            var forEachBlock = Quick.Block(parent.Get, "foreach( var CurItem in mCol )", new ICodeLine[]
            {
                Quick.Line("sAllKeyValues +=  CurItem.Key + KeyValueDelimiter + CurItem.Value + ItemDelimiter;"),
            });
            parent.Get.Children.Add(forEachBlock);
            
            parent.Get.Children.Add(Quick.Line(parent.Get, "CurItem = null;"));
            parent.Get.Children.Add(Quick.Line(parent.Get, "return sAllKeyValues;"));
            
            const string expected = 
@"        get
        {
            string sAllKeyValues;
            foreach( var CurItem in mCol )
            {
                sAllKeyValues +=  CurItem.Key + KeyValueDelimiter + CurItem.Value + ItemDelimiter;
            }
            CurItem = null;
            return sAllKeyValues;
        }
";
            var actual = parent.Get.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }

        public static CSharpPropertyPart QuickCSharpPropertyPart(ICodeLine parent, string line = null, PropertyPartType type = PropertyPartType.Get)
        {
            var target = new CSharpPropertyPart(parent, type)
            {
                Line = line,
                Encountered = true,
            };
            return target;
        }

        /*
        [TestMethod]
        public void NoChildrenWithLine()
        {
            var target = QuickCSharpProperty("F", "G", "H");
            
            const string expected = "    H\r\n    public G F { get; set; }\r\n";
            var actual = target.GenerateCode();

            Extensions.AreEqualFormatted(expected, actual);
        }


        [TestMethod]
        public void Comment_Get2Children_SetLine1Child()
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
            target.Get.Children.AddRange(
                new List<ICodeLine>()
                {
                    _.B(target.Get, "B"),
                    _.B(target.Get, "C")
                });

            target.Set.Line = "D";
            target.Set.Encountered = true;
            target.Set.Children.Add(_.B(target.Set, "E"));

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
            get.Children.Add(_.L(get, "Testing"));
            get.Children.Add(_.L(get, "123"));

            var actual = get.GenerateCode();
            var expected = "    get\r\n    {\r\n        Testing\r\n        123\r\n    }\r\n";
            Extensions.AreEqualFormatted(expected, actual);
        }
    */
    }
}