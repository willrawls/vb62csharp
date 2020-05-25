using System.Windows.Forms;

namespace MetX.VB6ToCSharp
{
    internal sealed class Debug
    {
        public static void WriteLine(string Message)
        {
            MessageBox.Show(Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}