using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mime;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CAssocItem_CodeShouldBeWellFormed
    {
        const string InputFilePath = @"CAssocItem.cls";

        [TestMethod]
        public void ProperlyTranslate_CAssocItem_cs()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var code = converter.GenerateCode(InputFilePath);

            // Must have
            var message = CheckConvertedCode.LookFor(code, true, new[]
            {
                "            return sGetToken",
                "            return m_sValue;",
                "m_",
                "        public string Key;",
                "    public string F",
                "public string Value",
            });

            // Must not have
            message += CheckConvertedCode.LookFor(code, false, new[]
            {
                "// freeware",
                "get;", "set;",
                "{;", "};",
                "';", "*;",
                "F = ",
                "assumed;",
                "risk.         //",
                "        ;",
                "delimiter.         //",
                "Ding ",
                "                    //",                           // Wrong indentation
                "        public class CAssocArray",                  // Wrong indentation
                "                public string Key;",               // Wrong indentation
                "                            public string Value",  // Wrong indentation
                "                    //  Retrieves the Nth",        // Wrong indentation
            });

            // One occurrence, no more, no less
            message += CheckConvertedCode.CheckOccurrences(code, 1, new[]
            {
                "return m_sValue;",
            });
            Assert.IsTrue(message == string.Empty, $"\n==========\n{message}\n==========\n{code}");

            Console.WriteLine(code);
        }

        [TestMethod]
        public void ByteForByteCorrectConversion_CAssocItem_cs()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var actual = converter.GenerateCode(InputFilePath);
            var expected = File.ReadAllText($"{InputFilePath}.expected");
            Assert.AreEqual(expected, actual);
        }
    }
}