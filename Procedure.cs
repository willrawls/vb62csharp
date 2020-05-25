using System.Collections;

namespace MetX.VB6ToCSharp
{
    public class Procedure
    {
        public string Comment { get; set; }

        public ArrayList LineList { get; set; }

        public ArrayList BottomLineList { get; set; }

        public string Name { get; set; }

        public ArrayList ParameterList { get; set; }

        public string ReturnType { get; set; }

        public string Scope { get; set; }

        public ProcedureType Type { get; set; }

        public Procedure()
        {
            LineList = new ArrayList();
            BottomLineList = new ArrayList();
            ParameterList = new ArrayList();
        }
    }
}