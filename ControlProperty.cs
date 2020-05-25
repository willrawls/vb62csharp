using System.Collections;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class ControlProperty
    {
        public string Comment { get; set; }
        public string Name { get; set; }

        public ArrayList PropertyList { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }

        public ControlProperty()
        {
            PropertyList = new ArrayList();
            Valid = false;
        }
    }
}