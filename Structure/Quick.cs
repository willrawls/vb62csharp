﻿using System.Collections.Generic;
using System.Linq;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

// ReSharper disable PossibleNullReferenceException

namespace MetX.VB6ToCSharp.Structure
{
    public static class Quick 
                          
    {
        /// <summary>
        /// Returns a new List<ICodeLine> adding up to two lines, setting the parent property on both, if provided
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static List<ICodeLine> List(ICodeLine parent = null, ICodeLine line1 = null, ICodeLine line2 = null)
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
        /// <returns>Children list with up to two lines in it</returns>
        public static List<ICodeLine> List(ICodeLine parent = null, string line1 = null, string line2 = null)
        {
            var lines = new List<ICodeLine>();
            
            if(parent != null && line1.IsNotEmpty())
            {
                lines.Add(Quick.Line(parent, line1));
            }

            if(parent != null && line2.IsNotEmpty())
            {
                lines.Add(Quick.Line(parent, line2));
            }

            return lines;
        }

        /// <summary>
        /// Builds a LineOfCode containing line
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line">The line of code to encapsulate</param>
        /// <returns></returns>
        public static LineOfCode Line(ICodeLine parent, string line)
        {
            var lineOfCode = new LineOfCode(parent, line);
            return lineOfCode;
        }

        /// <summary>
        /// Builds a LineOfCode containing line
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line">The line of code to encapsulate</param>
        /// <returns></returns>
        public static LineOfCode Line(string line)
        {
            var lineOfCode = new LineOfCode(new EmptyParent(), line);
            return lineOfCode;
        }

        /// <summary>
        /// Builds a Block with Line set and with 1 child code line
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line"></param>
        /// <param name="childLine"></param>
        /// <returns></returns>
        public static Block Block(ICodeLine parent, string line, ICodeLine childLine)
        {
            var block = new Block(parent, line);
            if (childLine == null) 
                return block;

            childLine.Parent = block;
            block.Children.Add(childLine);
            return block;
        }

        public static Block Block()
        {
            return Block(new EmptyParent(), null);
        }

        /// <summary>
        /// Builds a Block with Line set and with 1 child code line
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line"></param>
        /// <param name="childLine">Line added to Children</param>
        /// <param name="childLines"></param>
        /// <returns>The new Block</returns>
        public static Block Block(ICodeLine parent, string line, string childLines = null)
        {
            var block = new Block(parent, line);
            if (childLines.IsEmpty()) 
                return block;

            if (!childLines.Contains("\n"))
            {
                var child = new LineOfCode(block, childLines);
                block.Children.Add(child);
                return block;
            }

            foreach(var lineOfCode in 
                childLines.Lines().Trimmed()
                .Where(x => x.Length > 0))
            {
                block.Children.AddLine(block, lineOfCode);
                //block.Children.AddLines(block, childLineOfCode.Lines().SelectMany(l => l.Trim));
            }

            return block;
        }

        /// <summary>
        /// Builds a Block with Line set and with 1 child code line
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="line"></param>
        /// <param name="childLine">Line added to Children</param>
        /// <param name="childLineOfCode"></param>
        /// <returns>The new Block</returns>
        public static Block Block(ICodeLine parent, string line, ICodeLine[] children)
        {
            var block = new Block(parent, line);
            foreach(var child in children)
            {
                child.Parent = block;
                child.ResetIndent(block.Indent + 1);
                block.Children.Add(child);
            }   
            return block;
        }


        /// <summary>
        /// Used by unit tests to creat the top level parent but which does nothing
        /// </summary>
        /// <param name="startingIndent">-1 so the children start with a 0 indentation</param>
        /// <returns></returns>
        public static EmptyParent Top(int startingIndent = -1)
        {
            return new EmptyParent(startingIndent);
        }
    }
}