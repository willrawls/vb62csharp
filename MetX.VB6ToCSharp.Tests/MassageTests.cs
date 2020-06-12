using System;
using System.Collections.Generic;
using MetX.Scripts;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class MassageTests
    {
        const string inputFilePath = @"I:\OneDrive\data\code\Slice and Dice\Sandy\CAssocItem.cls";
        const string outputPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC";

        static List<string> NoNos = new List<string>
        {
            "get;", "set;", "{;", "};", "';"
        };

        [TestMethod]
        public void DetermineIfLineGetsASemicolonTest_PublicObjectFred_Brace()
        {
            var actual = Massage.DetermineIfLineGetsASemicolon("    public object Fred", "    {");

            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.Length > 0);

            Console.WriteLine(actual);
            Assert.IsFalse(actual.Contains("Fred;"));
            Assert.IsFalse(actual.Contains("{;"));
        }
    }
}