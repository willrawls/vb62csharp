using System.Windows.Forms;

namespace MetX.VB6ToCSharp
{
    public class Debug
    {
        public static void WriteLine(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}