using System;
using System.IO;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    public class CodeFragmentTestItem
    {
        public string ExpectedCSharpCode;
        public string VB6Code;

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

        public bool RunTest()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);

            var actual = converter
                .GenerateCodeFragment(VB6Code);

            Console.WriteLine(Environment.NewLine
                              + "---"
                              + Environment.NewLine
                              + VB6Code.GetRidOfEmptyLines()
                              + Environment.NewLine
                              + "---"
                              + Environment.NewLine);

            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.IsNotEmpty(), "actual is empty (no code returned)");

            Assert.AreEqual(
                ExpectedCSharpCode.GetRidOfEmptyLines(),
                actual.GetRidOfEmptyLines(), Environment.NewLine 
                                             + "---" 
                                             + VB6Code.GetRidOfEmptyLines());

            Console.WriteLine(Environment.NewLine
                              + actual.GetRidOfEmptyLines()
                              + Environment.NewLine
                              + "---");
            return true;
        }
    }
}