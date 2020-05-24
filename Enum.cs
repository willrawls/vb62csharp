using System.Collections;

namespace VB2C
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