using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    ///     Summary description for Control.
    /// </summary>
    public class Control : AbstractBlock
    {
        public bool Container;
        public bool InvisibleAtRuntime;
        public string Name;
        public string Owner;

        //public List<ControlProperty> PropertyList = new List<ControlProperty>();
        public string Type;

        public bool Valid;

        public Control(ICodeLine parent, string line = null) : base(parent, line)
        {
        }
    }
}