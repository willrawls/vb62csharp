using System.Text;

namespace MetX.VB6ToCSharp
{
    public static class Extensions
    {
        public static string Indent(this string originalLines, int indentLevel = 1)
        {
            if (indentLevel < 1)
                return originalLines;

            var result = new StringBuilder();
            var lines = originalLines.Replace("\r", "").Split('\n');
            var indentForLine = Tools.Indent(indentLevel);
            foreach (var line in lines)
            {
                result.AppendLine(indentForLine + line.Trim());
            }

            var code = result.ToString();
            return code;
        }
    }
}