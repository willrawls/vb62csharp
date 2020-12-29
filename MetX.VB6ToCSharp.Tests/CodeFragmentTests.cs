using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeFragmentTests
    {
        private string _folderPath = "./FragmentTests";

        [DataTestMethod]
        [DataRow("ConvertProperty_Simple")]
        public void TestLoadedFromFile(string testName)
        {
            var test = new CodeFragmentTestItem(Path.Combine(_folderPath, testName + ".fragmenttest"));
            test.RunTest();
        }
        
    }
}