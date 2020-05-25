using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Enum.
    /// </summary>
    public class Enum
    {
        public List<EnumItem> ItemList { get; set; }

        public string Name { get; set; }

        public string Scope { get; set; }

        public Enum()
        {
            ItemList = new List<EnumItem>();
        }
    }
}