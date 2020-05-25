using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Control.
    /// </summary>
    public class Control
    {
        public bool Container;
        public bool InvisibleAtRuntime;
        public string Name;

        public string Owner;
        public List<ControlProperty> PropertyList;
        public string Type;
        public bool Valid;

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