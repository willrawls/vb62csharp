using System.Windows.Forms;

namespace MetX.VB6ToCSharp
{
    internal sealed class Debug
    {
        public static void WriteLine(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}