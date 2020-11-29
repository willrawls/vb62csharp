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

            target.ResetIndent(1);
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
            
            target.ResetIndent(1);
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
            parent.ResetIndent(1);
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

    }
}