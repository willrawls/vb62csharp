using System;
using System.Collections.Generic;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class ModuleCodeShouldNotContainCertainStrings
    {
        const string inputFilePath = @"I:\OneDrive\data\code\Slice and Dice\Sandy\CAssocItem.cls";
        const string outputPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC";

        static List<string> NoNos = new List<string>
        {
            "get;", "set;", "{;", "};", "';"
        };

        [TestMethod]
        public void In_CAssocItem_cs()
        {
            ICodeLine parent = new EmptyParent(-1); 
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(parent, inputFilePath, outputPath);
            
            Assert.IsNotNull(code);
            Assert.IsTrue(code.Length > 0);

            var noNoList = "\nShould not contain: \n";
            foreach (var noNo in NoNos)
            {
                if(code.Contains(noNo))
                    noNoList += $"  {noNo}\n";
            }

            foreach (var noNo in NoNos)
            {
                Assert.IsFalse(code.Contains(noNo), noNoList);
            }
        }
    }
}