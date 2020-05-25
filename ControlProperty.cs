using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class ControlProperty
    {
        public string Comment;
        public string Name;

        public List<ControlProperty> PropertyList;
        public bool Valid;
        public string Value;

        public ControlProperty()
        {
            PropertyList = new List<ControlProperty>();
            Valid = false;
        }
    }
}