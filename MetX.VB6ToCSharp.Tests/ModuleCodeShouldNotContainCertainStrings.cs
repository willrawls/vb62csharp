using System;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CAssocArray_CodeShouldBeWellFormed
    {
        const string InputFilePath = "CAssocArray.cls";

        [TestMethod]
        public void ProperlyTranslate_CAssocArray_cs()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(InputFilePath);

            // Must have
            var message = CheckConvertedCode.LookFor(code, true, new[]
            {
                "public Dictionary<string,string> mCol;",
                @"foreach( var CurItem in mCol )
                {",
                "                    sAllKeyValues +=",
                "static string ",
                "Add(sKey, sValue);",
                "string sT;",
                "// EH_CAssocArray_Item_Continue:",
                "// goto EH_CAssocArray_Item_Continue;",
                "Item = Add(sIndexKey);",
                "long TokenCount;",
                "public CAssocArray()",
                "public ~CAssocArray()",
                "TreeToAll_AddChildren(sAll, CurChild);",
                "public void TreeToAll_AddChildren()",
                "foreach( var CurNode in tvwX.Nodes)",
                @"public void Remove()
    {
",
                "int",
                "NodeStack.Length",
                "",
            });

            // Must not have
            message += CheckConvertedCode.LookFor(code, false, new[]
            {
                "public unknown CurItem;",
                "public object CurItem;",
                "string static ",
                "string static sT;",
                "                sT.Contains(KeyValueDelimiter) > 0 )", // if( sT.Contains(KeyValueDelimiter) > 0 )
                "Class_Initialize",
                "Class_Terminate",
                "{;",
                @"        ;
        ;
",
                "foreach( var CurNode in tvwX.Nodes);",
                "Do Until",
                "        .Value.Substring",
                "Loop;",
                "Integer",
                "UBound(NodeStack)",
                @"        };
        };
        };
",
                "static",
                "",
                "",
                "",
            });

            // One occurrence, no more, no less
            message += CheckConvertedCode.CheckOccurrences(code, 1, new[]
            {
                "",
            });


            Assert.IsTrue(message == string.Empty, $"\n==========\n{message}\n==========\n{code}");

            Console.WriteLine(code);
        }
    }
}