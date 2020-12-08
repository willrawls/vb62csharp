using System.Collections.Generic;
using MetX.VB6ToCSharp.Interface;

namespace MetX.VB6ToCSharp.Structure
{
    public class CodeLineList : List<ICodeLine>
    {
        public CodeLineList AddLine(ICodeLine parent, string lineOfCode) 
        {
            var line = Quick.Line(parent, lineOfCode);
            Add(line);
            return this;
        }

        public CodeLineList AddLines(ICodeLine parent, IList<string> linesOfCode)
        {
            foreach (var lineOfCode in linesOfCode)
                AddLine(parent, lineOfCode);
            return this;
        }
    }
}