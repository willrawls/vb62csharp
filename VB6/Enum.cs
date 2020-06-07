using System.Collections.Generic;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    /// Summary description for Enum.
    /// </summary>
    public class Enum
    {
        public List<EnumItem> ItemList;

        public string Name;

        public string Scope;

        public Enum()
        {
            ItemList = new List<EnumItem>();
        }
    }
}