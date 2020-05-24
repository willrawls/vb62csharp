using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VB2C
{
    public static class Program
    {
        public static int Main(string[] args)
        {
            ConvertAll(@"I:\SandyB\");
            return 0;
        }

        public static void ConvertAll(string folderPath)
        {
            foreach (var fileToDelete in Directory.EnumerateFiles(folderPath))
                File.Delete(fileToDelete);

            var fileSets = new List<IEnumerable<string>>
            {
                Directory.EnumerateFiles(@"I:\Sandy", "*.cls"),
                Directory.EnumerateFiles(@"I:\Sandy", "*.frm"),
                Directory.EnumerateFiles(@"I:\Sandy", "*.bas"),
            };

            foreach (var fileSet in fileSets)
            foreach (var clsFile in fileSet)
            {
                var toConvert = new ConvertCode();
                toConvert.ParseFile(clsFile, folderPath);
            }
        }
    }
}