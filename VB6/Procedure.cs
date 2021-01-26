using System.Collections.Generic;
using System.Linq;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
{
    public class Procedure : Indentifier
    {
        public string Comment;

        public List<string> LineList;

        public List<string> BottomLineList;

        public string Name;

        public List<Parameter> ParameterList;

        public string ReturnType;

        public string Scope;

        public ProcedureType Type;

        public Procedure()
        {
            LineList = new List<string>();
            BottomLineList = new List<string>();
            ParameterList = new List<Parameter>();
        }

        public string GenerateCode()
        {
            var block = new Block
            {
                Line = Scope + " "
            };
            switch (Type)
            {
                case ProcedureType.ProcedureSub:
                    block.Line += "void ";
                    break;

                case ProcedureType.ProcedureFunction:
                    block.Line += ReturnType + " ";
                    break;

                case ProcedureType.ProcedureEvent:
                    block.Line += "void ";
                    break;
            }

            block.Line += Name;
            
            if (ParameterList.Count == 0)
            {
                block.Line += "()";
            }
            else
            {
                var firstParameter = true;
                foreach (var parameter in ParameterList)
                {
                    if (firstParameter)
                    {
                        firstParameter = false;
                    }
                    else
                    {
                        block.Line += ",";
                    }
                    block.Line += parameter.Type + " " + parameter.Name + " ";
                    if (parameter.IsOptional)
                        block.Line += "= " + parameter.OptionalValue;
                }
            }

            foreach (var line in LineList.Select(l => l.Trim()))
            {
                block.Children.AddLines(block.Parent, line.Lines().Trimmed());
            }

            foreach (var line in BottomLineList.Select(l => l.Trim()))
            {
                block.Children.AddLines(block.Parent, line.Lines().Trimmed());
            }
            return block.GenerateCode();
        }

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}