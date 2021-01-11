namespace MetX.VB6ToCSharp.VB6
{
    public class Parameter
    {
        public string Name;
        public string Pass;
        public string Type;
        public bool IsOptional;
        public string DefaultValue;

        public string OptionalValue
        {
            get
            {
                if (!IsOptional)
                    return "";
                return Type.ToLower() == "string" 
                    ? $"\"{DefaultValue}\" " 
                    : DefaultValue;
            }
        }
    }
}