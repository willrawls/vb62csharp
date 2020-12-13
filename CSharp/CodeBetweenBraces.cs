using System;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.RegularExpressions;

namespace MetX.VB6ToCSharp.CSharp
{
    public class CodeBetweenBraces : IComparable<CodeBetweenBraces>
    {
        public string StartingCode;
            
        public string BeforeOpenBrace;
        public string CodeFoundInsideBraces;
        public string AfterCloseBrace;

        public int IndexOfOpenBrace;
        public int IndexOfCloseBrace;
        public int StartedLookingAtIndex;
        public int IndexOfCode;
        public bool FindResult;

        public static Regex Regex =
            new Regex("{((?>[^{}]+|{(?<c>)|}(?<-c>))*(?(c)(?!)))",
                RegexOptions.Compiled);

        public CodeBetweenBraces() { }

        public CodeBetweenBraces(string startingCode, int startLookingAtIndex)
        {
            StartingCode = startingCode;
            StartedLookingAtIndex = startLookingAtIndex;

            var splits = CodeBetweenBraces.Regex.Split(startingCode).Where(s => s.Length > 0).ToArray();

                
            FindResult = splits.Length == 3;
            if (!FindResult) return;

            BeforeOpenBrace = splits[0] + "{";
            CodeFoundInsideBraces = splits[1];
            AfterCloseBrace = splits[2];
            IndexOfCode = CodeBetweenBraces.Regex.Match(startingCode, 0).Index;
            IndexOfOpenBrace = IndexOfCode - 1;
            IndexOfCloseBrace = startingCode.Length + IndexOfCode + 1;
        }

        public static CodeBetweenBraces Factory(string code, int startAtIndex = 0)
        {
            var result = new CodeBetweenBraces(code, startAtIndex);
            return result;
        }

        public override bool Equals(object obj) => Equals(obj as CodeBetweenBraces);

        public bool Equals(CodeBetweenBraces other)
        {
            return StartingCode == other.StartingCode && BeforeOpenBrace == other.BeforeOpenBrace && CodeFoundInsideBraces == other.CodeFoundInsideBraces && AfterCloseBrace == other.AfterCloseBrace && IndexOfOpenBrace == other.IndexOfOpenBrace && IndexOfCloseBrace == other.IndexOfCloseBrace && StartedLookingAtIndex == other.StartedLookingAtIndex && IndexOfCode == other.IndexOfCode && FindResult == other.FindResult;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (StartingCode != null ? StartingCode.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (BeforeOpenBrace != null ? BeforeOpenBrace.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (CodeFoundInsideBraces != null ? CodeFoundInsideBraces.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (AfterCloseBrace != null ? AfterCloseBrace.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ IndexOfOpenBrace;
                hashCode = (hashCode * 397) ^ IndexOfCloseBrace;
                hashCode = (hashCode * 397) ^ StartedLookingAtIndex;
                hashCode = (hashCode * 397) ^ IndexOfCode;
                hashCode = (hashCode * 397) ^ FindResult.GetHashCode();
                return hashCode;
            }
        }

        public int CompareTo(CodeBetweenBraces other)
        {
            if (ReferenceEquals(this, other)) return 0;
            if (ReferenceEquals(null, other)) return 1;
                
            var startingCodeComparison = string.Compare(StartingCode, other.StartingCode, StringComparison.InvariantCulture);
            if (startingCodeComparison != 0) return startingCodeComparison;
                
            var beforeOpenBraceComparison = string.Compare(BeforeOpenBrace, other.BeforeOpenBrace, StringComparison.InvariantCulture);
            if (beforeOpenBraceComparison != 0) return beforeOpenBraceComparison;
                
            var codeInsideBracesComparison = string.Compare(CodeFoundInsideBraces, other.CodeFoundInsideBraces, StringComparison.InvariantCulture);
            if (codeInsideBracesComparison != 0) return codeInsideBracesComparison;
                
            var afterCloseBraceComparison = string.Compare(AfterCloseBrace, other.AfterCloseBrace, StringComparison.InvariantCulture);
            if (afterCloseBraceComparison != 0) return afterCloseBraceComparison;
                
            var indexOfOpenBraceComparison = IndexOfOpenBrace.CompareTo(other.IndexOfOpenBrace);
            if (indexOfOpenBraceComparison != 0) return indexOfOpenBraceComparison;
                
            var indexOfCloseBraceComparison = IndexOfCloseBrace.CompareTo(other.IndexOfCloseBrace);
            if (indexOfCloseBraceComparison != 0) return indexOfCloseBraceComparison;
                
            var startedFindAtIndexComparison = StartedLookingAtIndex.CompareTo(other.StartedLookingAtIndex);
            if (startedFindAtIndexComparison != 0) return startedFindAtIndexComparison;
                
            var indexOfCodeComparison = IndexOfCode.CompareTo(other.IndexOfCode);
            if (indexOfCodeComparison != 0) return indexOfCodeComparison;

            return FindResult.CompareTo(other.FindResult);
        }

        public string Diff(CodeBetweenBraces other)
        {
            if (Equals(other))
                return "";

            var differences = new StringBuilder();

            if (ReferenceEquals(this, other)) return "";
            if (ReferenceEquals(null, other)) return "Other is null";

            differences.AppendLine("--- Differences found:");

            var comparison = string.Compare(StartingCode, other.StartingCode, StringComparison.InvariantCulture);
            if (comparison != 0)
                differences.AppendLine($"StartingCode [[{StartingCode}]] \\r\\n[[{other.StartingCode}]]\\r\\n");

            comparison = string.Compare(BeforeOpenBrace, other.BeforeOpenBrace, StringComparison.InvariantCulture);
            if (comparison != 0) 
                differences.AppendLine($"BeforeOpenBrace [[{BeforeOpenBrace}]] \\r\\n[[{other.BeforeOpenBrace}]]\\r\\n");

            comparison = string.Compare(CodeFoundInsideBraces, other.CodeFoundInsideBraces, StringComparison.InvariantCulture);
            if (comparison != 0)
                differences.AppendLine($"CodeFoundInsideBraces [[{CodeFoundInsideBraces}]] \\r\\n[[{other.CodeFoundInsideBraces}]]\\r\\n");

            comparison = string.Compare(AfterCloseBrace, other.AfterCloseBrace, StringComparison.InvariantCulture);
            if (comparison != 0)
                differences.AppendLine($"AfterCloseBrace [[{AfterCloseBrace}]] \\r\\n[[{other.AfterCloseBrace}]]\\r\\n");

            comparison = IndexOfOpenBrace.CompareTo(other.IndexOfOpenBrace);
            if (comparison != 0) 
                differences.AppendLine($"IndexOfOpenBrace [[{IndexOfOpenBrace}]] \\r\\n[[{other.IndexOfOpenBrace}]]\\r\\n");

            comparison = IndexOfCloseBrace.CompareTo(other.IndexOfCloseBrace);
            if (comparison != 0)
                differences.AppendLine($"IndexOfCloseBrace [[{IndexOfCloseBrace}]] \\r\\n[[{other.IndexOfCloseBrace}]]\\r\\n");

            comparison = StartedLookingAtIndex.CompareTo(other.StartedLookingAtIndex);
            if (comparison != 0)
                differences.AppendLine($"StartedLookingAtIndex [[{StartedLookingAtIndex}]] \\r\\n[[{other.StartedLookingAtIndex}]]\\r\\n");

            comparison = IndexOfCode.CompareTo(other.IndexOfCode);
            if (comparison != 0) 
                differences.AppendLine($"IndexOfCode [[{IndexOfCode}]] \\r\\n[[{other.IndexOfCode}]]\\r\\n");


            if (FindResult != other.FindResult)
                differences.AppendLine($"FindResult: [[{FindResult}]] [[{other.FindResult}]]");

            return differences.ToString();
        }
    }
}