using System;
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

        [DataTestMethod]
        [DataRow("public Dictionary<string,string> mCol;")]
        [DataRow("foreach(var CurItem in mCol) {")]
        [DataRow("                    sAllKeyValues +=")]
        [DataRow("static string ")]
        [DataRow("Add(sKey, sValue);")]
        [DataRow("string sT;")]
        [DataRow("// EH_CAssocArray_Item_Continue:")]
        [DataRow("// goto EH_CAssocArray_Item_Continue;")]
        [DataRow("Item = Add(sIndexKey);")]
        [DataRow("long TokenCount;")]
        [DataRow("public CAssocArray()")]
        [DataRow("public ~CAssocArray()")]
        [DataRow("TreeToAll_AddChildren(sAll, CurChild);")]
        [DataRow("public void TreeToAll_AddChildren()")]
        [DataRow("foreach( var CurNode in tvwX.Nodes)")]
        [DataRow("int CurSubItem;")]
        [DataRow("(CurNode.Tag)")]
        [DataRow("while(CurChild == null)")]
        [DataRow("TreeToAll_AddChildren(sAll, CurNode);")]
        public void MustContain(string textToFind)
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(InputFilePath);

            var message = code.MustHave(textToFind);
            Assert.IsTrue(message == string.Empty, $"\n==========\n{message}\n==========\n{code}");
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
        [DataRow("        .Value.Substring")]
        [DataRow("static CurSubItem /*As*/ int;")]
        [DataRow("(                                                                                        ")]
        [DataRow(@"        };
")]
        [DataRow(@"        ;
")]
        public void MustNotContain(string textToFind)
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(InputFilePath);

            var message = code.MustNotHave(textToFind);
            Assert.IsTrue(message == string.Empty, $"\n==========\n{message}\n==========\n{code}");
        }
    }
}