﻿using System;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;

// ReSharper disable InconsistentNaming

namespace MetX.VB6ToCSharp.Structure
{
    public class LineOfCode : Indentifier, ICodeLine
    {
        public int Indent => Parent.Indent + 1;

        public string Line { get; set; }
        public ICodeLine Parent { get; set; }

        public LineOfCode(ICodeLine parent, string line = null)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            Line = line;
            _internalIndent = Indent;
        }

        public LineOfCode(string line)
        {
            Line = line;
        }

        public virtual string GenerateCode()
        {
            return Indentation + (Line ?? "") + "\r\n";
        }

        public virtual bool IsEmpty() => Line.IsEmpty();

        public virtual bool IsNotEmpty() => Line.IsNotEmpty();
    }
}