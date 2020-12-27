using System.Linq;
using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CodeFragmentTests
    {
        [TestMethod]
        public void ConvertProperty_Simple()
        {
            var vb6Code = 
@"Public Property Let Example(sNewValue As String)
    m_Example = sNewValue
End Property

Public Property Get Example() As String
Attribute Example.VB_UserMemId = 0
    Example = m_Example
End Property";
            
            var expectedCSharpCode = @"
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


";
            TestCodeFragmentConversion(vb6Code, expectedCSharpCode);
        }
        
        public void TestCodeFragmentConversion(string vb6Code, string expectedCSharpCode)
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            
            var actual = converter.GenerateCodeFragment(vb6Code);
            
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.IsNotEmpty());
            Assert.AreEqual(expectedCSharpCode, actual);
        }
    }
}