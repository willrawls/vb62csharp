using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeFragmentTests
    {
        private string _folderPath = "./FragmentTests";

        [DataTestMethod]
        [DataRow("Add")]
        [DataRow("ConvertProperty_Simple")]
        [DataRow("DoWhileIfElse")]
        [DataRow("ForEach2")]
        [DataRow("ForEachWith1")]
        [DataRow("IfElse")]
        [DataRow("ItemEqualsAdd")]
        [DataRow("OnErrorResume")]
        [DataRow("StaticSubVariable")]
        [DataRow("SubCall")]
        [DataRow("VariableDeclarations")]
        [DataRow("With.Tag")]
        public void RunTestFromFile(string testName)
        {
            var test = new CodeFragmentTestItem(
                Path.Combine(_folderPath, testName + ".fragmenttest"));
            test.RunTest();
        }
        
    }
}