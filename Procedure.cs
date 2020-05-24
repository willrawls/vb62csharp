using System.Collections;

namespace VB2C
{
    /// <summary>
    /// Summary description for Procedure.
    /// </summary>
    ///

    public enum ProcedureType
    {
        ProcedureEvent = 1,
        ProcedureSub = 2,
        ProcedureFunction = 3
    }

    public class Procedure
    {
        public string Comment { get; set; }

        public ArrayList LineList { get; set; }

        public string Name { get; set; }

        public ArrayList ParameterList { get; set; }

        public string ReturnType { get; set; }

        public string Scope { get; set; }

        public ProcedureType Type { get; set; }

        public Procedure()
        {
            LineList = new ArrayList();
            ParameterList = new ArrayList();
        }
    }
}