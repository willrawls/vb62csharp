﻿using System.IO;
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
        [DataRow("DoWhile1")]
        [DataRow("DoWhileIfElse")]
        [DataRow("ForEach2")]
        [DataRow("ForEachWith1")]
        [DataRow("IfElse")]
        [DataRow("ItemEqualsAdd")]
        [DataRow("OnErrorResume")]
        [DataRow("ProperIndentation1")]
        [DataRow("StaticSubVariable")]
        [DataRow("SubCall")]
        [DataRow("VariableDeclarations")]
        [DataRow("With.Tag")]
        public void RunTestFromFile(string testName)
        {
            RunCodeFragmentTest(testName);
        }

        [TestMethod]
        public void RunJustOneTestFromFile()
        {
            RunCodeFragmentTest("DoWhile1");
        }

        private void RunCodeFragmentTest(string testName)
        {
            var filePath = Path.Combine(_folderPath, testName + ".fragmenttest");
            var test = new CodeFragmentTestItem(filePath);
            test.RunTest();
        }
    }
}