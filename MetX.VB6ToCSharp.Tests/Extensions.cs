using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MetX.VB6ToCSharp.Tests
{
    public static class Extensions
    {
        public static void AreEqualFormatted(object expected, object actual, string message = null)
        {
            expected = "\n" + expected;
            actual = "\n" + actual;
            Assert.AreEqual(expected, actual, message);
            Console.WriteLine("Expected:" + expected);
            Console.WriteLine("Actual  :" +  actual);
        }
    }
}