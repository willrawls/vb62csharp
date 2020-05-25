using System.Collections;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Enum.
    /// </summary>
    public class Enum
    {
        public ArrayList ItemList { get; set; }

        public string Name { get; set; }

        public string Scope { get; set; }

        public Enum()
        {
            ItemList = new ArrayList();
        }
    }
}