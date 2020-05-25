using System;
using System.IO;
using System.Windows.Forms;

namespace MetX.VB6ToCSharp
{
    public class App
    {
        public const string ConfigOutPath = "OutPath";
        public const string ConfigSetting = "Setting";
        public static XmlConfig Config;
        public static FrmConvert MainForm;

        // configuration constants
        public const string ConfigFile = "vb2c.xml";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            // get current directory
            // index 0 contain path and name of exe file
            var sBinPath = Path.GetDirectoryName(args[0]);

            var showGui = false;
            if (args.Length > 1)
                showGui = args[1].ToLower().Replace("-", "") == "gui";

            // create configuration object
            Config = new XmlConfig(sBinPath + Path.DirectorySeparatorChar + ConfigFile);

            if (showGui)
            {
                // create main screen
                MainForm = new FrmConvert();
                MainForm.Show();
                Application.Run(MainForm);
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }
}