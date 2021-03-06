﻿using MetX.Library;
using MetX.VB6ToCSharp.VB6;

namespace MetX.VB6ToCSharp.CSharp
{
    public abstract class Indentifier
    {
        protected string _indentation;
        protected string _secondIndentation;
        protected int _internalIndent;

        public string Indentation
        {
            get
            {
                if(_indentation.IsEmpty())
                    _indentation = Tools.Indent(_internalIndent);
                return _indentation;
            }
        }

        public string SecondIndentation
        {
            get
            {
                if(_secondIndentation.IsEmpty())
                    _secondIndentation = Tools.Indent(_internalIndent + 1);
                return _secondIndentation;
            }
        }

        public int Indent
        {
            get => _internalIndent;
            set => _internalIndent = value;
        }

        public abstract void ResetIndent(int indentLevel);
    }
}