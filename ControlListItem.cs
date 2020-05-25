namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for ControlListItem.
    /// </summary>
    public class ControlListItem
    {
        public string CsharpName { get; set; }
        public bool InvisibleAtRuntime { get; set; }
        public bool Unsupported { get; set; }
        public string Vb6Name { get; set; }
    }
}