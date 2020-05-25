using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class Property
    {
        public string Comment;

        public string Direction;

        public List<string> LineList;

        public string Name;

        public List<Parameter> ParameterList;

        public string Scope;

        public string Type;

        public Property()
        {
            LineList = new List<string>();
            ParameterList = new List<Parameter>();
        }
    }
}