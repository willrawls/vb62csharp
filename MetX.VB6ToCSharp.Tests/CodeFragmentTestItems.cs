using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MetX.VB6ToCSharp.Tests
{
    public class CodeFragmentTestItems : List<CodeFragmentTestItem>
    {
        public static CodeFragmentTestItems Load(string folderPath)
        {
            var result = new CodeFragmentTestItems();
            result.AddRange(
                Directory.EnumerateFiles(
                        folderPath, "*.fragmenttest")
                    .Select(file => 
                        new CodeFragmentTestItem(file)));

            return result;
        }
    }
}