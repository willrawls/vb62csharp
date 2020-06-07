using System.Collections.Generic;
using System.IO;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp
{
    public static class Program
    {
        public const string BaseFolder = @"I:\OneDrive\data\code\Slice and Dice";
        public static string SourceCodeFolderPath = $@"{BaseFolder}Sandy\";
        public static string OutputFolderPath = $@"{BaseFolder}\SandyC\";
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
                Directory.EnumerateFiles(SourceCodeFolderPath, "*.bas"),
                Directory.EnumerateFiles(SourceCodeFolderPath, "*.frm"),
            };

            // -1 lets the generated code start at the beginning of a line
            ICodeLine parent = new FirstParent(-1); 

            foreach (var fileSet in fileSets)
            foreach (var file in fileSet)
            {
                new ModuleConverter(parent)
                    .ConvertFile(parent, file, OutputFolderPath);
            }
        }
    }
}