using System.Text.RegularExpressions;
using MetX.VB6ToCSharp.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class OneLineComplexTests
    {
        [TestMethod]
        public void InstrTest()
        {
            var data = OneLineComplex.Shuffle[0];
            var actual = data.Regex.Replace("string z = Instr(x,y)", data.ReplacePattern ?? "");
            Assert.AreEqual("x.Contains(y)", actual);
        }

        [TestMethod]
        public void XequalsXcode()
        {
            var regex = new Regex(@"(.*) = (\1\b)(.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var actual = regex.Replace("Fred = Fred + George && Frank", "$1 += $3");
            Assert.AreEqual("\nFred += George && Frank", "\n" + actual);
        }
        // Fred += George && Frank
        // Fred += Fred  + George && Frank
    }
}