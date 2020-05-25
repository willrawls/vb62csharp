using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public class Procedure
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
    }
}