using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Property.
    /// </summary>
    public class Property : IAmAProperty
    {
        public string Direction;
        public List<string> LineList;

        public string Comment { get; set; }
        public string Name { get; set; }
        public bool Valid { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }

        public Property()
        {
            LineList = new List<string>();
        }

        public string GenerateTargetCode()
        {
            throw new System.NotImplementedException();
        }
    }
}