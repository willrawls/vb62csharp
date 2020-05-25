using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    public class Procedure
    {
        public string Comment { get; set; }

        public List<string> LineList { get; set; }

        public List<string> BottomLineList { get; set; }

        public string Name { get; set; }

        public List<Parameter> ParameterList { get; set; }

        public string ReturnType { get; set; }

        public string Scope { get; set; }

        public ProcedureType Type { get; set; }

        public Procedure()
        {
            LineList = new List<string>();
            BottomLineList = new List<string>();
            ParameterList = new List<Parameter>();
        }
    }
}