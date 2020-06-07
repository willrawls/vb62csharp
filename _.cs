using System.Collections.Generic;
using MetX.Library;
// ReSharper disable PossibleNullReferenceException

namespace MetX.VB6ToCSharp
{
    public static class _    // abcdefghijklmnopqrstuvwxyz
                             // ABCDEFGHIJKLMNOPQRSTUVWXYZ
    {
        /// <summary>
        /// Returns a new List<ICodeLine> adding up to two lines, setting the parent property on both, if provided
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static List<ICodeLine> C(ICodeLine parent = null, ICodeLine line1 = null, ICodeLine line2 = null)
        {
            var lines = new List<ICodeLine>();
            
            if(line1.IsNotEmpty())
            {
                if(parent != null)
                    line1.Parent = parent;
                lines.Add(line1);
            }

            if (line2.IsEmpty()) 
                return lines;

            if(parent != null)
                line2.Parent = parent;
            lines.Add(line2);

            return lines;
        }

        /// <summary>
        /// Returns a new List<ICodeLine> adding up to two lines
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static List<ICodeLine> C(ICodeLine parent = null, string line1 = null, string line2 = null)
        {
            var lines = new List<ICodeLine>();
            
            if(parent != null && line1.IsNotEmpty())
            {
                lines.Add(_.L(parent, line1));
            }

            if(parent != null && line2.IsNotEmpty())
            {
                lines.Add(_.L(parent, line2));
            }

            return lines;
        }

        public static LineOfCode L(ICodeLine parent, string line)
        {
            var lineOfCode = new LineOfCode(parent, line);
            return lineOfCode;
        }

        public static Block B(ICodeLine parent, string line, ICodeLine childLine)
        {
            var block = new Block(parent, line);
            if (childLine == null) 
                return block;

            childLine.Parent = block;
            block.Children.Add(childLine);
            return block;
        }
        
        public static Block B(ICodeLine parent, string line, string childLineOfCode = null)
        {
            var block = new Block(parent, line);
            if (childLineOfCode.IsEmpty()) 
                return block;

            var child = new LineOfCode(block, childLineOfCode);
            block.Children.Add(child);
            return block;
        }

        /*
        public static IBlock X(this IBlock parent, string line = "", string childLineOfCode = null)
        {
            return B(parent,
        }
        */

        public static EmptyParent E(int startingIndent = -1)
        {
            return new EmptyParent(startingIndent);
        }
    }
}