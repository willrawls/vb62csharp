using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public class CSharpPropertyPart
    {
        public bool Encountered;
        public List<string> LineList = new List<string>();
        public List<Parameter> ParameterList = new List<Parameter>();

        public string GenerateCode()
        {
            return string.Empty;
        }
    }
}