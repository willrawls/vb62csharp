using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.VB6ToCSharp.CSharp;

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
            var result = new StringBuilder();

            // public void WriteResX ( List<string> mImageList, string OutPath, string ModuleName )
            result.Append(Indentation + Scope + " ");
            switch (Type)
            {
                case ProcedureType.ProcedureSub:
                    result.Append("void");
                    break;

                case ProcedureType.ProcedureFunction:
                    result.Append(ReturnType);
                    break;

                case ProcedureType.ProcedureEvent:
                    result.Append("void");
                    break;
            }

            // name
            result.Append(" " + Name);
            // parameters
            if (ParameterList.Count == 0)
            {
                result.AppendLine("()");
            }

            // start body
            result.AppendLine(Indentation + "{");

            foreach (var line in LineList.Select(l => l.Trim()))
                if (line.Length > 0)
                    result.AppendLine($"{SecondIndentation}{line};");
                else
                    result.AppendLine();

            foreach (var line in BottomLineList.Select(l => l.Trim()))
                if (line.Length > 0)
                    result.AppendLine($"{SecondIndentation}{line};");
                else
                    result.AppendLine();

            // end procedure
            result.AppendLine(Indentation + "}");
            return result.ToString();
        }

        public override void ResetIndent(int indentLevel)
        {
            _internalIndent = indentLevel;
            _indentation = null;
            _secondIndentation = null;
        }
    }
}