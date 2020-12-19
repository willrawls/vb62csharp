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
            //var code = converter.GenerateCode(InputFilePath);
            
            var code = codeForMustContain.IsEmpty() 
                ? converter.GenerateCode(InputFilePath) 
                : codeForMustContain;

            if (codeForMustContain.IsEmpty())
                codeForMustContain = code;

            var message = code.MustHave(textToFind);
            if (message.IsNotEmpty())
            {
                code = Isolate(code, textToFind, 100);
            }

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
        [DataRow("        .Value.Substring")]
        [DataRow("static CurSubItem /*As*/ int;")]
        [DataRow("(                                                                                        ")]
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
                isolatedCode = Isolate(code, textToFind, 100);

            Assert.IsTrue(message.IsEmpty(), $"{message}{isolatedCode}");
            Console.WriteLine($"{code}");
        }

        private string Isolate(string code, string textToFind, int beforeAndAfter)
        {
            var indexOfTextToFind =
                code.IndexOf(textToFind, StringComparison.InvariantCultureIgnoreCase);

            var startIndex = indexOfTextToFind - beforeAndAfter;
            var endIndex = indexOfTextToFind + beforeAndAfter + textToFind.Length;

            if (startIndex < 0)
                startIndex = 0;
            
            if(endIndex > code.Length)
                if (startIndex == 0)
                    return code;
                else
                    return "... " + code.Substring(startIndex).Trim();

            return "... " + code.Substring(startIndex, endIndex - startIndex).Trim() + " ...";
        }

        private static string codeForMustContain = null;
        private static string codeForMustNotContain = null;
    }
}