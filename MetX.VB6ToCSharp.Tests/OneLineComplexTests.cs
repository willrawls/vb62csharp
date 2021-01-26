using MetX.VB6ToCSharp.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class OneLineComplexTests
    {
        [TestMethod] public void NoChangeTest()
        {
            var actual = RegExReplacements.Shuffle("string z = (x,y)");
            Assert.AreEqual("string z = (x,y)", actual);
        }

        [TestMethod] public void fromXasYto_Y_X()
        {
            var actual = RegExReplacements.Shuffle("fred As henry");
            Assert.AreEqual("henry fred;", actual);
        }

        [TestMethod] public void InstrTest()
        {
            var actual = RegExReplacements.Shuffle("string z = Instr(x,y)");
            Assert.AreEqual("x.Contains(y)", actual);
        }

        // x = x + y...
        [TestMethod]
        public void XequalsXcode()
        {
            var actual = RegExReplacements.Shuffle("Fred = Fred + George && Frank");
            Assert.AreEqual("\nFred +=  George && Frank;", "\n" + actual);
        }

        // For x = y To z
        [TestMethod]
        public void ForXEqualsYtoZ()
        {
            var actual = RegExReplacements.Shuffle("For xxx = yyy To ZZZ");
            Assert.AreEqual("\nfor(var xxx = yyy; xxx < ZZZ; xxx++) //SOB//", "\n" + actual);
        }

        // AddItem x
        [TestMethod]
        public void AddItem()
        {
            var actual = RegExReplacements.Shuffle("AddItem x");
            Assert.AreEqual("\n" + "AddItem(x)", "\n" + actual);
        }

        [TestMethod]
        public void from_AddX_to_Add_x_()
        {
            var actual = RegExReplacements.Shuffle("Add x,y,z");
            Assert.AreEqual("\n" + "Add(x,y,z)", "\n" + actual);
            
            actual = RegExReplacements.Shuffle("Add x,y");
            Assert.AreEqual("\n" + "Add(x,y)", "\n" + actual);
            
            actual = RegExReplacements.Shuffle("Add x");
            Assert.AreEqual("\n" + "Add(x)", "\n" + actual);
        }

        [TestMethod]
        public void from_SetListIndexX_to_SetListIndex_x_()
        {
            var actual = RegExReplacements.Shuffle("SetListIndex x");
            Assert.AreEqual("\n" + "SetListIndex(x)", "\n" + actual);
        }

        // Mid$(x,y)
        [TestMethod]
        public void Mid()
        {
            var actual = RegExReplacements.Shuffle("Mid$( x, y)");
            Assert.AreEqual("\n" + " x.Substring( y)", "\n" + actual);

            actual = RegExReplacements.Shuffle("Mid( x, y)");
            Assert.AreEqual("\n" + " x.Substring( y)", "\n" + actual);
        }

        // UCase$(x)
        [TestMethod]
        public void UCase()
        {
            var actual = RegExReplacements.Shuffle("UCase$(x)");
            Assert.AreEqual("\nx.ToUpper()", "\n" + actual);

            actual = RegExReplacements.Shuffle("UCase( x)");
            Assert.AreEqual("\n" + " x.ToUpper()", "\n" + actual);
        }

        [TestMethod]
        public void Left()
        {
            var actual = RegExReplacements.Shuffle("left( x,y)");
            Assert.AreEqual("\nx.Substring(0, y)", "\n" + actual);

            actual = RegExReplacements.Shuffle("left$( x,y)");
            Assert.AreEqual("\nx.Substring(0, y)", "\n" + actual);
        }

        [TestMethod]
        public void Right()
        {
            var actual = RegExReplacements.Shuffle("right(x,y)");
            Assert.AreEqual("\nx.Substring(x.Length - y)", "\n" + actual);

            actual = RegExReplacements.Shuffle(@"right$( x,y)");
            Assert.AreEqual("\nx.Substring(x.Length - y)", "\n" + actual);
        }

        // Do While x > y
        [TestMethod]
        public void DoWhile()
        {
            var actual = RegExReplacements.Shuffle("Do While x > y");
            Assert.AreEqual("\n" + "while(x > y)//SOB//", "\n" + actual);
        }
    }
}