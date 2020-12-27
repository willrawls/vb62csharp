using MetX.Library;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class ModuleConverterTests
    {
        [TestMethod]
        public void ConvertPropertyFromACodeFragment()
        {
            ICodeLine parent = new EmptyParent(-1);
            var converter = new ModuleConverter(parent);
            var codeIn = 
@"Public Property Let Example(sNewValue As String)
    m_sValue = sNewValue
End Property

Public Property Get Example() As String
Attribute Value.VB_UserMemId = 0
    Value = m_Example
End Property";
            var expected = @"
        public string Example
        {
            get
            {
                return m_Example;
            }

            set
            {
                m_sValue = value;
            }

        }
";

            var actual = converter.GenerateCodeFragment(codeIn);
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.IsNotEmpty());
            Assert.AreEqual(expected, actual);
        }
    }
}