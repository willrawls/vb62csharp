using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class ControlProperty
    {
        public string Comment { get; set; }
        public string Name { get; set; }

        public List<ControlProperty> PropertyList { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }

        public ControlProperty()
        {
            PropertyList = new List<ControlProperty>();
            Valid = false;
        }
    }
}