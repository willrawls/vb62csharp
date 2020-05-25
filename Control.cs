using System.Collections;

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
        public ArrayList PropertyList { get; set; }
        public string Type { get; set; }
        public bool Valid { get; set; }

        public Control()
        {
            PropertyList = new ArrayList();
        }

        public void PropertyAdd(ControlProperty oProperty)
        {
            PropertyList.Add(oProperty);
        }
    }
}