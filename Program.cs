using System.Collections.Generic;
using System.IO;

namespace MetX.VB6ToCSharp
{
    public static class Program
    {
        public const string SourceCodeFolderPath = @"I:\OneDrive\data\code\Slice and Dice\Sandy\";
        public const string OutputFolderPath = @"I:\OneDrive\data\code\Slice and Dice\SandyC\";
        public const bool ClearOutputFolder = false;

        public static int Main(string[] args)
        {
            ConvertAllFiles();
            return 0;
        }

        public static void ConvertAllFiles()
        {
            if(ClearOutputFolder)
            {
                // Delete previous run
                foreach (var fileToDelete in Directory.EnumerateFiles(OutputFolderPath))
                {
                    File.SetAttributes(fileToDelete, FileAttributes.Normal);
                    File.Delete(fileToDelete);
                }
            }

            var fileSets = new List<IEnumerable<string>>
            {
                Directory.EnumerateFiles(SourceCodeFolderPath, "*.cls"),
                Directory.EnumerateFiles(SourceCodeFolderPath, "*.frm"),
                Directory.EnumerateFiles(SourceCodeFolderPath, "*.bas"),
            };

            foreach (var fileSet in fileSets)
            foreach (var file in fileSet)
            {
                new ModuleConverter()
                    .ConvertFile(file, OutputFolderPath);
            }
        }
    }
}