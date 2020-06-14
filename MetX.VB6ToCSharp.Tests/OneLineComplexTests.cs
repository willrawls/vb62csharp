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
            var actual = OneLineComplex.Shuffle("string z = (x,y)");
            Assert.AreEqual("string z = (x,y)", actual);
        }

        [TestMethod] public void fromXasYto_Y_X()
        {
            var actual = OneLineComplex.Shuffle("fred As henry");
            Assert.AreEqual("henry fred;", actual);
        }

        [TestMethod] public void InstrTest()
        {
            var actual = OneLineComplex.Shuffle("string z = Instr(x,y)");
            Assert.AreEqual("x.Contains(y)", actual);
        }

        // x = x + y...
        [TestMethod]
        public void XequalsXcode()
        {
            var actual = OneLineComplex.Shuffle("Fred = Fred + George && Frank");
            Assert.AreEqual("\nFred +=  George && Frank;", "\n" + actual);
        }

        // For x = y To z
        [TestMethod]
        public void ForXEqualsYtoZ()
        {
            var actual = OneLineComplex.Shuffle("For xxx = yyy To ZZZ");
            Assert.AreEqual("\nfor(var xxx = yyy; xxx < ZZZ; xxx++) //SOB//", "\n" + actual);
        }

        // AddItem x
        [TestMethod]
        public void AddItem()
        {
            var actual = OneLineComplex.Shuffle("AddItem x");
            Assert.AreEqual("\n" + "AddItem(x)", "\n" + actual);
        }

        [TestMethod]
        public void from_AddX_to_Add_x_()
        {
            var actual = OneLineComplex.Shuffle("Add x,y,z");
            Assert.AreEqual("\n" + "Add(x,y,z)", "\n" + actual);
            
            actual = OneLineComplex.Shuffle("Add x,y");
            Assert.AreEqual("\n" + "Add(x,y)", "\n" + actual);
            
            actual = OneLineComplex.Shuffle("Add x");
            Assert.AreEqual("\n" + "Add(x)", "\n" + actual);
        }

        [TestMethod]
        public void from_SetListIndexX_to_SetListIndex_x_()
        {
            var actual = OneLineComplex.Shuffle("SetListIndex x");
            Assert.AreEqual("\n" + "SetListIndex(x)", "\n" + actual);
        }

        // Mid$(x,y)
        [TestMethod]
        public void Mid()
        {
            var actual = OneLineComplex.Shuffle("Mid$( x, y)");
            Assert.AreEqual("\n" + " x.Substring( y)", "\n" + actual);

            actual = OneLineComplex.Shuffle("Mid( x, y)");
            Assert.AreEqual("\n" + " x.Substring( y)", "\n" + actual);
        }

        // UCase$(x)
        [TestMethod]
        public void UCase()
        {
            var actual = OneLineComplex.Shuffle("UCase$(x)");
            Assert.AreEqual("\nx.ToUpper()", "\n" + actual);

            actual = OneLineComplex.Shuffle("UCase( x)");
            Assert.AreEqual("\n" + " x.ToUpper()", "\n" + actual);
        }

        [TestMethod]
        public void Left()
        {
            var actual = OneLineComplex.Shuffle("left( x,y)");
            Assert.AreEqual("\nx.Substring(0, y)", "\n" + actual);

            actual = OneLineComplex.Shuffle("left$( x,y)");
            Assert.AreEqual("\nx.Substring(0, y)", "\n" + actual);
        }

        [TestMethod]
        public void Right()
        {
            var actual = OneLineComplex.Shuffle("right(x,y)");
            Assert.AreEqual("\nx.Substring(x.Length - y)", "\n" + actual);

            actual = OneLineComplex.Shuffle(@"right$( x,y)");
            Assert.AreEqual("\nx.Substring(x.Length - y)", "\n" + actual);
        }

        // Do While x > y
        [TestMethod]
        public void DoWhile()
        {
            var actual = OneLineComplex.Shuffle("Do While x > y");
            Assert.AreEqual("\n" + "while(x > y)//SOB//", "\n" + actual);
        }
    }
}