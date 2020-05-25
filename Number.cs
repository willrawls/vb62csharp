namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Tools.
    /// </summary>
    public class Number
    {
        public static bool IsOdd(int iNumber)
        {
            if (iNumber != (iNumber / 2) * 2)
            { return true; }
            else
            { return false; }
        }
    }
}