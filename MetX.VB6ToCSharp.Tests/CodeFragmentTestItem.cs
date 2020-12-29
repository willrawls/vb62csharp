using System;
using System.IO;
using System.Runtime.CompilerServices;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    public class CodeFragmentTestItem
    {
        public string VB6Code;
        public string ExpectedCSharpCode;

        public CodeFragmentTestItem(string filePath)
        {
            var delimiter = "~~~~";
            var contents = File.ReadAllText(filePath);
            
            VB6Code = contents.TokenAt(1, delimiter);
            ExpectedCSharpCode = contents.TokenAt(2, delimiter);
        }

        public CodeFragmentTestItem()
        {
            
        }

        public void RunTest()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            
            var actual = converter.GenerateCodeFragment(VB6Code);
            
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.IsNotEmpty(), "actual is empty (no code returned)");
            Assert.AreEqual(ExpectedCSharpCode, actual);
            Console.WriteLine(actual);
        }
    }
}