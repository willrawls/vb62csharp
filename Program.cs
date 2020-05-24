using System.IO;
using System.Linq;

namespace VB2C
{
    public static class Program
    {
        public static void ConvertAll(string folderPath)
        {
            foreach (var fileToDelete in Directory.EnumerateFiles(folderPath))
                File.Delete(fileToDelete);

            var clsFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.cls").ToList();
            var frmFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.frm").ToList();
            var basFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.bas").ToList();

            foreach (var fileSet in new[] { clsFiles, basFiles, frmFiles })
                foreach (var clsFile in fileSet)
                {
                    var ConvertObject = new ConvertCode();
                    ConvertObject.ParseFile(clsFile, folderPath);
                }
        }

        public static int Main(string[] args)
        {
            ConvertAll(@"I:\SandyB\");

            return 0;
        }
    }
}