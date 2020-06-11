using System.Text.RegularExpressions;
using MetX.VB6ToCSharp.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class OneLineComplexTests
    {
        [TestMethod] public void NoChangeTest()
        {
            var actual = OneLineComplex.Shuffle[0].Replace("string z = (x,y)");
            Assert.AreEqual("string z = (x,y)", actual);
        }

        [TestMethod] public void InstrTest()
        {
            var actual = OneLineComplex.Shuffle[0].Replace("string z = Instr(x,y)");
            Assert.AreEqual("x.Contains(y)", actual);
        }

        // x = x + y...
        [TestMethod]
        public void XequalsXcode()
        {
            var actual = OneLineComplex.Shuffle[1].Replace("Fred = Fred + George && Frank");
            Assert.AreEqual("\nFred +=  George && Frank;", "\n" + actual);
        }

        // For x = y To z
        [TestMethod]
        public void ForXEqualsYtoZ()
        {
            var actual = OneLineComplex.Shuffle[2].Replace("For xxx = yyy To ZZZ");
            Assert.AreEqual("\nfor(var xxx = yyy; xxx < ZZZ; xxx++) //SOB//", "\n" + actual);
        }

        // Add x
        [TestMethod]
        public void Add()
        {
            var actual = OneLineComplex.Shuffle[3].Replace("Add x,y,z");
            Assert.AreEqual("\n" + "Add(x,y,z)", "\n" + actual);
        }

        // Mid$(x,y)
        [TestMethod]
        public void Mid()
        {
            var actual = OneLineComplex.Shuffle[4].Replace("Mid$( x, y)");
            Assert.AreEqual("\n" + " x.Substring( y)", "\n" + actual);
        }

        // Do While x > y
        [TestMethod]
        public void DoWhile()
        {
            var actual = OneLineComplex.Shuffle[5].Replace("Do While x > y");
            Assert.AreEqual("\n" + "while(x > y)//SOB//", "\n" + actual);
        }
    }
}