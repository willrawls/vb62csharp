using System.Collections;

namespace VB2C
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class Property
    {
        public string Comment { get; set; }

        public string Direction { get; set; }

        public ArrayList LineList { get; set; }

        public string Name { get; set; }

        public ArrayList ParameterList { get; set; }

        public string Scope { get; set; }

        public string Type { get; set; }

        public Property()
        {
            LineList = new ArrayList();
            ParameterList = new ArrayList();
        }
    }
}