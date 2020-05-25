using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class Property
    {
        public string Comment { get; set; }

        public string Direction { get; set; }

        public List<string> LineList { get; set; }

        public string Name { get; set; }

        public List<Parameter> ParameterList { get; set; }

        public string Scope { get; set; }

        public string Type { get; set; }

        public Property()
        {
            LineList = new List<string>();
            ParameterList = new List<Parameter>();
        }
    }
}