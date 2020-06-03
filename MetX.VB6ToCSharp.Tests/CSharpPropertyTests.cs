using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    [TestClass]
    public class CSharpPropertyTests
    {
        [TestMethod]
        public void SimpleTest()
        {
            var target = new CSharpProperty(1);
            target.Comment = "' TheComment";
            target.Name = "F";
            target.Type = "G";

            target.Get = new CSharpPropertyPart(target, PropertyPartType.Get);
            target.Get.Line = "A";
            target.Get.Children = new List<AbstractCodeBlock>();
            target.Get.Children.Add(new CodeBlock(target.Get, "B"));
            target.Get.Children.Add(new CodeBlock(target.Get, "C"));
            target.Get.Encountered = true;

            target.Set = new CSharpPropertyPart(target, PropertyPartType.Set);
            target.Set.Line = "D";
            target.Set.Children = new List<AbstractCodeBlock>();
            target.Set.Children.Add(new CodeBlock(target.Get, "E"));
            target.Set.Encountered = true;

            var expected =
@"    // TheComment
    public G F
    {
        get
        {
            A
            {
                B
                C
            }
        }
        set
        {
            D
            {
                E
            }
        }
";
            var actual = target.GenerateCode();
            Console.WriteLine(actual);

            Assert.AreEqual(expected, actual);
        }
    }
}