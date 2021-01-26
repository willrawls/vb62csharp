using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class ExtensionTests
    {
        [TestMethod]
        public void Deflate_Simple()
        {
            Assert.AreEqual("abc", "abc".Deflate());
            Assert.AreEqual("\"abc", "\"abc".Deflate());
            Assert.AreEqual("\"abc\"", "\"abc\"".Deflate());
            Assert.AreEqual("abc\"", "abc\"".Deflate());
            Assert.AreEqual("a b c", "a  b  c".Deflate());
            Assert.AreEqual("a\"b\"c", "a\"b\"c".Deflate());
            Assert.AreEqual("a \"  b  \" c", "  a  \"  b  \"  c  ".Deflate());
        }

        
        [TestMethod]
        public void Swap_Simple()
        {
            Assert.AreEqual("", "".PutABeforeBOnce("a", "b"));
            Assert.AreEqual("", ((string)null).PutABeforeBOnce("a", "b"));
            Assert.AreEqual("1234", "1234".PutABeforeBOnce("a", "b"));

            Assert.AreEqual("12", "21".PutABeforeBOnce("1", "2"));
            Assert.AreEqual("1 2", "2 1".PutABeforeBOnce("1", "2"));

            Assert.AreEqual("bca", "abc".PutABeforeBOnce("bc", "a"));

            Assert.AreEqual("abc", "abc".PutABeforeBOnce("ac", "b"));
            Assert.AreEqual("acb", "bac".PutABeforeBOnce("ac", "b"));
            Assert.AreEqual("a cb", "a cb".PutABeforeBOnce("ac", "b"));

            Assert.AreEqual("ab+a=b", "ba+a=b".PutABeforeBOnce("a", "b"));

            Assert.AreEqual("baaa", "aaab".PutABeforeBOnce("b", "a")); // a already before b
            Assert.AreEqual("babb", "abbb".PutABeforeBOnce("b", "a"));

            Assert.AreEqual("ba-a-b", "ab-a-b".PutABeforeBOnce("b", "a"));
            Assert.AreEqual("..b--a==a", "..a--b==a".PutABeforeBOnce("b", "a"));
        }

        [TestMethod]
        public void ExamineAndAdjust_CSharpProperty_Get_NoAdjustment()
        {
            var root = new EmptyParent();
            var property = new CSharpProperty(root);

            property.Get.Line = "x = y;";
            property.Get.Children.Add(Quick.Line(property, "z = a;"));
            property.Get.LineListAfter.Add(Quick.Line(property, "b =c;"));

            property.Get.ExamineAndAdjust();
            Assert.AreEqual("x = y;", property.Get.Line);
            Assert.AreEqual("z = a;", property.Get.Children[0].Line);
            Assert.AreEqual("b =c;", property.Get.LineListAfter[0].Line);
        }

        // --- in class static

        [TestMethod]
        public void ExamineAndAdjustLine_InClass_Static()
        {
            Assert.AreEqual("static string x;", "string static x;".ExamineAndAdjustLine(CurrentlyInArea.Class));
            Assert.AreEqual("    static int b = 12;", "    int static b = 12;".ExamineAndAdjustLine(CurrentlyInArea.Class));
            Assert.AreEqual("static Collection c;", "Collection static c;".ExamineAndAdjustLine(CurrentlyInArea.Class));
        }

        [TestMethod]
        public void ExamineAndAdjust_InClass_Static()
        {
            var root = new EmptyParent();
            var module = new Module(root);
            module.VariableList.Add(new Variable
            {
                Name = "Fred",
                Scope = "public",
                Type = "string static",
            });

            module.ExamineAndAdjust();
            Assert.AreEqual("static string", module.VariableList[0].Type);
        }

        // ---- in property static
        [TestMethod]
        public void ExamineAndAdjustLine_InProperty_Static()
        {
            Assert.AreEqual("string x;", "string static x;".ExamineAndAdjustLine(CurrentlyInArea.Property));
            Assert.AreEqual("int b = 12;", "int static b = 12;".ExamineAndAdjustLine(CurrentlyInArea.Property));
            Assert.AreEqual("Collection c;", "Collection static c;".ExamineAndAdjustLine(CurrentlyInArea.Property));
        }

        [TestMethod]
        public void ExamineAndAdjust_CSharpProperty_Get_Static()
        {
            var root = new EmptyParent();
            var property = new CSharpProperty(root);

            property.Get.Line = "string static x;";
            property.Get.Children.Add(Quick.Line(property, "int static b = 12;"));
            property.Get.LineListAfter.Add(Quick.Line(property, "Collection static c;"));

            property.Get.ExamineAndAdjust();
            Assert.AreEqual("string x;", property.Get.Line);
            Assert.AreEqual("int b = 12;", property.Get.Children[0].Line);
            Assert.AreEqual("Collection c;", property.Get.LineListAfter[0].Line);
        }
    }
}