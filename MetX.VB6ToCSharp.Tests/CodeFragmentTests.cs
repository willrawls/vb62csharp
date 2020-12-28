using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeFragmentTests
    {
        [TestInitialize]
        public void TestInitialize()
        {
            Items = CodeFragmentTestItems.Load("FragmentTests");

        }

        public CodeFragmentTestItems Items { get; set; }

        [TestMethod]
        public void Test1()
        {
            Items[0].RunTest();
        }
        
        [TestMethod]
        public void ConvertProperty_Simple()
        {
            var item = new CodeFragmentTestItem
            {
                VB6Code =
                    @"Public Property Let Example(sNewValue As String)
    m_Example = sNewValue
End Property

Public Property Get Example() As String
Attribute Example.VB_UserMemId = 0
    Example = m_Example
End Property",
                ExpectedCSharpCode = @"
        public string Example
        {
            get
            {
                return m_Example;
            }

            set
            {
                m_Example = value;
            }

        }


",
            };
            item.RunTest();
        }
        
        /*
        public void TestCodeFragmentConversion(string vb6Code, string expectedCSharpCode)
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            
            var actual = converter.GenerateCodeFragment(vb6Code);
            
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.IsNotEmpty());
            Assert.AreEqual(expectedCSharpCode, actual);
        }
    */
    }
}