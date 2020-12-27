using System;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class ProperlyTranslateCAssocArray_cls
    {
        private const string InputFilePath = "CAssocArray.cls";

        private static string codeForMustContain;
        private static string codeForMustNotContain;

        [DataTestMethod]
        [DataRow("                    sAllKeyValues +=")]
        [DataRow("(CurNode.Tag)")]
        [DataRow("EH_CAssocArray_Item_Continue:")]
        [DataRow("// goto EH_CAssocArray_Item_Continue;")]
        [DataRow("Add(sKey, sValue);")]
        [DataRow("foreach( var CurNode in tvwX.Nodes)")]
        [DataRow("foreach(var CurItem in mCol) {")]
        [DataRow("int CurSubItem;")]
        [DataRow("Item = Add(sIndexKey);")]
        [DataRow("long TokenCount;")]
        [DataRow("public ~CAssocArray()")]
        [DataRow("public CAssocArray()")]
        [DataRow("public Dictionary<string,string> mCol;")]
        [DataRow("public void TreeToAll_AddChildren()")]
        [DataRow("static string ")]
        [DataRow("string sT;")]
        [DataRow("TreeToAll_AddChildren(sAll, CurChild);")]
        [DataRow("TreeToAll_AddChildren(sAll, CurNode);")]
        [DataRow("while(CurChild == null)")]
        public void MustContain(string textToFind)
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            //var code = converter.GenerateCode(InputFilePath);

            var code = codeForMustContain.IsEmpty()
                ? converter.GenerateCode(InputFilePath)
                : codeForMustContain;

            if (codeForMustContain.IsEmpty())
                codeForMustContain = code;

            var message = code.MustHave(textToFind);
            if (message.IsNotEmpty()) code = CSharp.Extensions.Isolate(code, textToFind, 100);

            Assert.IsTrue(message == string.Empty, $"{message}\n==========\n{code}");
            //Console.WriteLine($"{code}");
        }

        [DataTestMethod]
        [DataRow("public unknown CurItem;")]
        [DataRow("public object CurItem;")]
        [DataRow("string static ")]
        [DataRow("string static sT;")]
        [DataRow("                sT.Contains(KeyValueDelimiter) > 0 )")]
        [DataRow("Class_Initialize")]
        [DataRow("Class_Terminate")]
        [DataRow("{;")]
        [DataRow("foreach( var CurNode in tvwX.Nodes);")]
        [DataRow("Do Until")]
        [DataRow("        .Value.Substring")]
        [DataRow("Loop;")]
        [DataRow("Integer")]
        [DataRow("UBound(NodeStack)")]
        [DataRow("static")]
        [DataRow("Is null")]
        [DataRow("(.Tag)")]
        [DataRow("static CurSubItem /*As*/ int;")]
        [DataRow(
            "(                                                                                        ")]
        [DataRow(@"        };
")]
        [DataRow(@"        ;
")]
        public void MustNotContain(string textToFind)
        {
            //if (!textToFind.Contains("Class_"))
            //    return;

            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);

            var code = codeForMustNotContain.IsEmpty()
                ? converter.GenerateCode(InputFilePath)
                : codeForMustNotContain;

            if (codeForMustNotContain.IsEmpty())
                codeForMustNotContain = code;

            var message = code.MustNotHave(textToFind);

            var isolatedCode = code;
            if (message.IsNotEmpty())
                isolatedCode = CSharp.Extensions.Isolate(code, textToFind, 100);

            Assert.IsTrue(message.IsEmpty(), $"{message}{isolatedCode}");
            Console.WriteLine($"{code}");
        }

        [DataTestMethod]
        [DataRow("string static Fred;", "static", "string", "static string Fred;" )]
        [DataRow("long static Fred;", "static", "long", "static long Fred;" )]
        [DataRow(" int    static   Fred; ", "static", "int", " static    int   Fred; " )]
        public void PutABeforeBOnce_Simple(string target, string a, string b, string expected)
        {
            var actual = CSharp.Extensions.PutABeforeBOnce(target, a, b);
            Assert.AreEqual(expected, actual, $"{target} | {expected} | {actual}");
        }
    }
}