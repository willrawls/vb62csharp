using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class FrmConvert : Form
    {
        private Button cmdConvert;
        private Button cmdExit;
        private Button cmdLoad;
        private readonly IContainer components = null;
        private Button convertAllButton;
        private Label label;
        private Label label1;
        private string mFileName;
        private string mOutPath;
        private TextBox txtCSharp;
        private TextBox txtOutPath;
        private TextBox txtVb6;

        public FrmConvert()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>

        private void cmdConvert_Click(object sender, EventArgs e)
        {
            if (mFileName.Trim() != string.Empty)
            {
                // parse file
                var convertObject = new ConvertCode();
                if (txtOutPath.Text.Trim() == string.Empty)
                {
                    MessageBox.Show("Fill out path !", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!Directory.Exists(txtOutPath.Text.Trim()))
                {
                    MessageBox.Show("Out path not exists !", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                mOutPath = txtOutPath.Text;
                if (mOutPath.Substring(mOutPath.Length - 1, 1) != @"\")
                {
                    mOutPath = mOutPath + @"\";
                }
                convertObject.ParseFile(mFileName, mOutPath);

                // show result
                txtCSharp.Text = convertObject.Code;
            }
        }

        private void cmdExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void cmdLoad_Click(object sender, EventArgs e)
        {
            mFileName = FileOpen();
            if (mFileName != null)
            {
                // show content of file
                var reader = File.OpenText(mFileName);
                txtVb6.Text = reader.ReadToEnd();
                reader.Close();
            }
        }

        private void ConvertAllButton_Click(object sender, EventArgs e)
        {
            foreach (var fileToDelete in Directory.EnumerateFiles(@"I:\SandyB\"))
                File.Delete(fileToDelete);

            var clsFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.cls").ToList();
            var frmFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.frm").ToList();
            var basFiles = Directory.EnumerateFiles(@"I:\Sandy", "*.bas").ToList();

            foreach (var fileSet in new[] { clsFiles, basFiles, frmFiles })
                foreach (var clsFile in fileSet)
                {
                    var convertObject = new ConvertCode();
                    convertObject.ParseFile(clsFile, @"I:\SandyB\");

                    // show result
                    //txtCSharp.Text = ConvertObject.OutSourceCode;
                }
        }

        private string FileOpen()
        {
            var sFilter = "VB6 form (*.frm)|*.frm|VB6 module (*.bas)|*.bas|VB6 class (*.cls)|*.cls|All files (*.*)|*.*";
            string sResult = null;

            var oDialog = new OpenFileDialog();
            oDialog.Filter = sFilter;
            if (oDialog.ShowDialog() != DialogResult.Cancel)
            {
                sResult = oDialog.FileName;
            }
            return sResult;
        }

        private void frmConvert_Closing(object sender, CancelEventArgs e)
        {
            App.Config.WriteString(App.ConfigSetting, App.ConfigOutPath, txtOutPath.Text);
        }

        private void frmConvert_Load(object sender, EventArgs e)
        {
            txtOutPath.Text = App.Config.ReadString(App.ConfigSetting, App.ConfigOutPath, "");
            ConvertAllButton_Click(sender, e);
        }

        private void frmConvert_Resize(object sender, EventArgs e)
        {
            //-16
            txtVb6.Top = 8;
            txtVb6.Left = 8;
            txtVb6.Height = Height / 2 - 60;
            txtVb6.Width = Width - 24;

            txtCSharp.Left = 8;
            txtCSharp.Top = Height / 2 - 16;
            txtCSharp.Height = Height / 2 - 60;
            txtCSharp.Width = Width - 24;
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmConvert));
            this.cmdExit = new System.Windows.Forms.Button();
            this.label = new System.Windows.Forms.Label();
            this.txtCSharp = new System.Windows.Forms.TextBox();
            this.cmdConvert = new System.Windows.Forms.Button();
            this.cmdLoad = new System.Windows.Forms.Button();
            this.txtVb6 = new System.Windows.Forms.TextBox();
            this.txtOutPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.convertAllButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // cmdExit
            //
            this.cmdExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdExit.BackColor = System.Drawing.SystemColors.ControlLight;
            this.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdExit.Location = new System.Drawing.Point(616, 520);
            this.cmdExit.Name = "cmdExit";
            this.cmdExit.Size = new System.Drawing.Size(80, 40);
            this.cmdExit.TabIndex = 0;
            this.cmdExit.Text = "Exit";
            this.cmdExit.UseVisualStyleBackColor = false;
            this.cmdExit.Click += new System.EventHandler(this.cmdExit_Click);
            //
            // label
            //
            this.label.Location = new System.Drawing.Point(40, 16);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(128, 24);
            this.label.TabIndex = 0;
            this.label.Text = "label";
            //
            // txtCSharp
            //
            this.txtCSharp.AcceptsReturn = true;
            this.txtCSharp.AcceptsTab = true;
            this.txtCSharp.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtCSharp.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCSharp.Location = new System.Drawing.Point(8, 264);
            this.txtCSharp.MaxLength = 327670;
            this.txtCSharp.Multiline = true;
            this.txtCSharp.Name = "txtCSharp";
            this.txtCSharp.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCSharp.Size = new System.Drawing.Size(688, 248);
            this.txtCSharp.TabIndex = 3;
            this.txtCSharp.WordWrap = false;
            //
            // cmdConvert
            //
            this.cmdConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdConvert.Location = new System.Drawing.Point(100, 520);
            this.cmdConvert.Name = "cmdConvert";
            this.cmdConvert.Size = new System.Drawing.Size(80, 40);
            this.cmdConvert.TabIndex = 6;
            this.cmdConvert.Text = "Convert";
            this.cmdConvert.Click += new System.EventHandler(this.cmdConvert_Click);
            //
            // cmdLoad
            //
            this.cmdLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdLoad.Location = new System.Drawing.Point(8, 520);
            this.cmdLoad.Name = "cmdLoad";
            this.cmdLoad.Size = new System.Drawing.Size(80, 40);
            this.cmdLoad.TabIndex = 4;
            this.cmdLoad.Text = "Load";
            this.cmdLoad.Click += new System.EventHandler(this.cmdLoad_Click);
            //
            // txtVB6
            //
            this.txtVb6.AcceptsReturn = true;
            this.txtVb6.AcceptsTab = true;
            this.txtVb6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtVb6.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtVb6.Location = new System.Drawing.Point(8, 8);
            this.txtVb6.MaxLength = 327670;
            this.txtVb6.Multiline = true;
            this.txtVb6.Name = "txtVb6";
            this.txtVb6.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtVb6.Size = new System.Drawing.Size(688, 248);
            this.txtVb6.TabIndex = 2;
            this.txtVb6.WordWrap = false;
            //
            // txtOutPath
            //
            this.txtOutPath.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.txtOutPath.Location = new System.Drawing.Point(255, 530);
            this.txtOutPath.Name = "txtOutPath";
            this.txtOutPath.Size = new System.Drawing.Size(220, 20);
            this.txtOutPath.TabIndex = 7;
            //
            // label1
            //
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(190, 530);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Out path";
            //
            // ConvertAllButton
            //
            this.convertAllButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.convertAllButton.Location = new System.Drawing.Point(491, 520);
            this.convertAllButton.Name = "convertAllButton";
            this.convertAllButton.Size = new System.Drawing.Size(80, 40);
            this.convertAllButton.TabIndex = 6;
            this.convertAllButton.Text = "Convert all";
            this.convertAllButton.Click += new System.EventHandler(this.ConvertAllButton_Click);
            //
            // frmConvert
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.cmdExit;
            this.ClientSize = new System.Drawing.Size(704, 565);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtOutPath);
            this.Controls.Add(this.convertAllButton);
            this.Controls.Add(this.cmdConvert);
            this.Controls.Add(this.cmdLoad);
            this.Controls.Add(this.txtCSharp);
            this.Controls.Add(this.txtVb6);
            this.Controls.Add(this.cmdExit);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(712, 592);
            this.Name = "FrmConvert";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert VB6 to C#";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmConvert_Closing);
            this.Load += new System.EventHandler(this.frmConvert_Load);
            this.Resize += new System.EventHandler(this.frmConvert_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        //    private string FileSave()
        //    {
        //      string sFilter = "C# Files (*.cs)|*.cs" ;
        //      string sResult = null;
        //
        //      SaveFileDialog oDialog = new SaveFileDialog();
        //      oDialog.Filter = sFilter;
        //      if(oDialog.ShowDialog() != DialogResult.Cancel)
        //      {
        //        sResult = oDialog.FileName;
        //      }
        //      return sResult;
        //    }
    }
}