namespace MetX.VB6ToCSharp.Interface
{
    public interface IGenerate
    {
        string Final { get; }

        string GenerateCode();
    }
}