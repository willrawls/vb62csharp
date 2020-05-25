using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Control.
    /// </summary>
    public class Control
    {
        public bool Container { get; set; }
        public bool InvisibleAtRuntime { get; set; }
        public string Name { get; set; }

        public string Owner { get; set; }
        public List<ControlProperty> PropertyList { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }

        public Control()
        {
            PropertyList = new List<ControlProperty>();
        }

        public void PropertyAdd(ControlProperty property)
        {
            PropertyList.Add(property);
        }
    }
}