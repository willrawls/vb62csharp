using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Convert.
    /// </summary>
    public class ConvertCode
    {
        public const string ClassFirstLine = "1.0 CLASS";
        public const string FormFirstLine = "VERSION 5.00";
        public const string Indent2 = "  ";
        public const string Indent4 = "    ";
        public const string Indent6 = "      ";
        public const string ModuleFirstLine = "ATTRIBUTE";
        public readonly List<string> moOwnerStock;
        public VbFileType mFileType;
        public Module mSourceModule;
        public Module mTargetModule;
        public string ActionResult { get; set; }

        public string Code { get; set; }

        public string ProjectNamespace { get; set; } = "MetX.SliceAndDice";

        public ConvertCode()
        {
            moOwnerStock = new List<string>();
        }

        public void ConvertFormCode(string outPath, StringBuilder oResult)
        {
            // list of controls
            foreach (Control control in mTargetModule.ControlList)
            {
                if (!control.Valid)
                {
                    oResult.Append("//");
                }

                oResult.Append(Indent2 + " public System.Windows.Forms." + control.Type + " " + control.Name + ";\r\n");
            }

            oResult.Append(Indent4 + "/// <summary>\r\n");
            oResult.Append(Indent4 + "/// Required designer variable.\r\n");
            oResult.Append(Indent4 + "/// </summary>\r\n");
            oResult.Append(Indent4 + "public System.ComponentModel.Container components = null;\r\n");
            oResult.Append("\r\n");
            oResult.Append(Indent4 + "public " + mSourceModule.Name + "()\r\n");
            oResult.Append(Indent4 + "{\r\n");
            oResult.Append(Indent6 + "// Required for Windows Form Designer support\r\n");
            oResult.Append(Indent6 + "InitializeComponent();\r\n");
            oResult.Append("\r\n");
            oResult.Append(Indent6 + "// TODO: Add any constructor code after InitializeComponent call\r\n");
            oResult.Append(Indent4 + "}\r\n");

            oResult.Append(Indent4 + "/// <summary>\r\n");
            oResult.Append(Indent4 + "/// Clean up any resources being used.\r\n");
            oResult.Append(Indent4 + "/// </summary>\r\n");
            oResult.Append(Indent4 + "protected override void Dispose( bool disposing )\r\n");
            oResult.Append(Indent4 + "{\r\n");
            oResult.Append(Indent6 + "if( disposing )\r\n");
            oResult.Append(Indent6 + "{\r\n");
            oResult.Append(Indent6 + "  if (components != null)\r\n");
            oResult.Append(Indent6 + "  {\r\n");
            oResult.Append(Indent6 + "    components.Dispose();\r\n");
            oResult.Append(Indent6 + "  }\r\n");
            oResult.Append(Indent6 + "}\r\n");
            oResult.Append(Indent6 + "base.Dispose( disposing );\r\n");
            oResult.Append(Indent4 + "}\r\n");

            oResult.Append(Indent4 + "#region Windows Form Designer generated code\r\n");
            oResult.Append(Indent4 + "/// <summary>\r\n");
            oResult.Append(Indent4 + "/// Required method for Designer support - do not modify\r\n");
            oResult.Append(Indent4 + "/// the contents of this method with the code editor.\r\n");
            oResult.Append(Indent4 + "/// </summary>\r\n");
            oResult.Append(Indent4 + "public void InitializeComponent()\r\n");
            oResult.Append(Indent4 + "{\r\n");

            // if form contain images
            if (mTargetModule.ImagesUsed)
            {
                // System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
                oResult.Append(Indent6 + "System.Resources.ResourceManager resources = " +
                               "new System.Resources.ResourceManager(typeof(" + mTargetModule.Name + "));\r\n");
            }

            foreach (Control control in mTargetModule.ControlList)
            {
                if (!control.Valid)
                {
                    oResult.Append("//");
                }

                oResult.Append(Indent6 + "this." + control.Name
                               + " = new System.Windows.Forms." + control.Type
                               + "();\r\n");
            }

            // SuspendLayout part
            oResult.Append(Indent6 + "this.SuspendLayout();\r\n");
            // this.Frame1.ResumeLayout(false);
            // resume layout for each container
            foreach (Control control in mTargetModule.ControlList)
            {
                // check if control is container
                // !! for menu controls
                if (control.Container && control.Type != "MenuItem" && control.Type != "MainMenu")
                {
                    if (!control.Valid)
                    {
                        oResult.Append("//");
                    }

                    oResult.Append(Indent6 + "this." + control.Name + ".SuspendLayout();\r\n");
                }
            }

            // each controls and his property
            foreach (Control control in mTargetModule.ControlList)
            {
                oResult.Append(Indent6 + "//\r\n");
                oResult.Append(Indent6 + "// " + control.Name + "\r\n");
                oResult.Append(Indent6 + "//\r\n");

                // unsupported control
                if (!control.Valid)
                {
                    oResult.Append("/*");
                }

                // ImageList, Timer, Menu has't name property
                if ((control.Type != "ImageList") && (control.Type != "Timer")
                                                   && (control.Type != "MenuItem") && (control.Type != "MainMenu"))
                {
                    // control name
                    oResult.Append(Indent6 + "this." + control.Name + ".Name = "
                                   + (char)34 + control.Name + (char)34 + ";\r\n");
                }

                // write properties
                foreach (ControlProperty property in control.PropertyList)
                {
                    GetPropertyRow(oResult, control.Type, control.Name, property, outPath);
                }

                // if control is container for other controls
                var temp = string.Empty;
                foreach (Control oControl1 in mTargetModule.ControlList)
                {
                    // all controls ownered by current control
                    if ((oControl1.Owner == control.Name) && (!oControl1.InvisibleAtRuntime))
                    {
                        temp += Indent6 + Indent6 + "this." + oControl1.Name + ",\r\n";
                    }
                }

                if (temp != string.Empty)
                {
                    // exception for menu controls
                    if (control.Type == "MainMenu" || control.Type == "MenuItem")
                    {
                        // this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                        oResult.Append(Indent6 + "this." + control.Name
                                       + ".MenuItems.AddRange(new System.Windows.Forms.MenuItem[]\r\n");
                    }
                    else
                    {
                        // this. + control.Name + .Controls.AddRange(new System.Windows.Forms.Control[]
                        oResult.Append(Indent6 + "this." + control.Name
                                       + ".Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
                    }

                    oResult.Append(Indent6 + "{\r\n");
                    oResult.Append((string)temp);
                    // remove last comma, keep CRLF
                    oResult.Remove(oResult.Length - 3, 1);
                    // close addrange part
                    oResult.Append(Indent6 + "});\r\n");
                }

                // unsupported control
                if (!control.Valid)
                {
                    oResult.Append("*/");
                }
            }

            oResult.Append(Indent6 + "//\r\n");
            oResult.Append(Indent6 + "// " + mSourceModule.Name + "\r\n");
            oResult.Append(Indent6 + "//\r\n");
            oResult.Append(Indent6 + "this.Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
            oResult.Append(Indent6 + "{\r\n");

            // add control range to form
            foreach (Control control in mTargetModule.ControlList)
            {
                if (!control.Valid)
                {
                    oResult.Append("//");
                }

                // all controls ownered by main form
                if ((control.Owner == mSourceModule.Name) && (!control.InvisibleAtRuntime))
                {
                    oResult.Append(Indent6 + "      this." + control.Name + ",\r\n");
                }
            }

            // remove last comma, keep CRLF
            oResult.Remove(oResult.Length - 3, 1);
            // close addrange part
            oResult.Append(Indent6 + "});\r\n");

            // form name
            oResult.Append(Indent6 + "this.Name = " + (char)34 + mTargetModule.Name + (char)34 + ";\r\n");
            // exception for menu
            // this.Menu = this.mainMenu1;
            if (mTargetModule.MenuUsed)
            {
                foreach (Control control in mTargetModule.ControlList)
                {
                    if (control.Type == "MainMenu")
                    {
                        oResult.Append(Indent6 + "      this.Menu = " + control.Name + ";\r\n");
                    }
                }
            }

            // form properties
            foreach (ControlProperty property in mTargetModule.FormPropertyList)
            {
                if (!property.Valid)
                {
                    oResult.Append("//");
                }

                GetPropertyRow(oResult, mTargetModule.Type, "", property, outPath);
            }

            // resume layout for each container
            foreach (Control control in mTargetModule.ControlList)
            {
                // check if control is container
                if ((control.Container) && !(control.Type == "MenuItem") && !(control.Type == "MainMenu"))
                {
                    if (!control.Valid)
                    {
                        oResult.Append("//");
                    }

                    oResult.Append(Indent6 + "this." + control.Name + ".ResumeLayout(false);\r\n");
                }
            }

            // form
            oResult.Append(Indent6 + "this.ResumeLayout(false);\r\n");

            oResult.Append(Indent4 + "}\r\n");
            oResult.Append(Indent4 + "#endregion\r\n");
        }

        public bool ParseFile(string filename, string outputPath)
        {
            var version = string.Empty;
            bool result;

            // try recognize source code type depend by file extension
            var extension = filename.Substring(filename.Length - 3, 3);
            switch (extension.ToUpper())
            {
                case "FRM":
                    mFileType = VbFileType.VbFileForm;
                    break;

                case "BAS":
                    mFileType = VbFileType.VbFileModule;
                    break;

                case "CLS":
                    mFileType = VbFileType.VbFileClass;
                    break;

                default:
                    mFileType = VbFileType.VbFileUnknown;
                    break;
            }

            // open file
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                using (var reader = new StreamReader(stream))
                {
                    reader.BaseStream.Seek(0, SeekOrigin.Begin);
                    var line = reader.ReadLine() ?? string.Empty;
                    // verify type of file based on first line - form, module, class

                    // get first word from first line
                    var position = 0;
                    var temp = GetWord(line, ref position);
                    switch (temp.ToUpper())
                    {
                        // module first line
                        // 'Attribute VB_Name = "ModuleName"'
                        case ModuleFirstLine:
                            mFileType = VbFileType.VbFileModule;
                            break;

                        // form or class first line
                        // 'VERSION 5.00' or 'VERSION 1.0 CLASS'
                        case "VERSION":
                            position++;
                            version = GetWord(line, ref position);

                            mFileType = line.Contains(ClassFirstLine)
                                ? VbFileType.VbFileClass
                                : VbFileType.VbFileForm;
                            break;

                        default:
                            mFileType = VbFileType.VbFileUnknown;
                            break;
                    }

                    // if file is still unknown
                    if (mFileType == VbFileType.VbFileUnknown)
                    {
                        ActionResult = "Unknown file type";
                        return false;
                    }

                    mSourceModule = new Module
                    {
                        Version = version ?? "1.0",
                        FileName = filename
                    };

                    // now parse specifics of each type
                    switch (extension.ToUpper())
                    {
                        case "FRM":
                            mSourceModule.Type = "form";
                            result = ParseForm(reader);
                            break;

                        case "BAS":
                            mSourceModule.Type = "module";
                            result = ParseModule(reader);
                            break;

                        case "CLS":
                            mSourceModule.Type = "class";
                            result = ParseClass(reader);
                            break;
                    }

                    // parse remain - variables, functions, procedures
                    result = ParseProcedures(reader);

                    stream.Close();
                    reader.Close();
                }
            }

            // generate output file
            Code = GetCode(outputPath);

            // save result
            var outFileName = outputPath + mTargetModule.FileName;
            File.WriteAllText(outFileName, Code);

            // generate resx file if source form contain any images
            if ((mTargetModule.ImagesUsed))
            {
                WriteResX(mTargetModule.ImageList, outputPath, mTargetModule.Name);
            }

            return result;
        }

        public void ConvertEnumCode(StringBuilder oResult)
        {
            oResult.Append("\r\n");
            foreach (Enum enumItems in mTargetModule.EnumList)
            {
                // public enum VB_FILE_TYPE
                oResult.Append(Indent4 + enumItems.Scope + " enum " + enumItems.Name + "\r\n");
                oResult.Append(Indent4 + "{\r\n");

                foreach (EnumItem enumItem in enumItems.ItemList)
                {
                    // name
                    oResult.Append(Indent6 + enumItem.Name);

                    if (enumItem.Value != string.Empty)
                    {
                        oResult.Append(" = " + enumItem.Value);
                    }

                    // enum items delimiter
                    oResult.Append(",\r\n");
                }

                // remove last comma, keep CRLF
                oResult.Remove(oResult.Length - 3, 1);
                // end enum
                oResult.Append(Indent4 + "};\r\n");
            }
        }

        public void ConvertProcedureCode(StringBuilder oResult)
        {
            oResult.Append("\r\n");
            foreach (Procedure procedure in mTargetModule.ProcedureList)
            {
                // public void WriteResX ( List<string> mImageList, string OutPath, string ModuleName )
                oResult.Append(Indent4 + procedure.Scope + " ");
                switch (procedure.Type)
                {
                    case ProcedureType.ProcedureSub:
                        oResult.Append("void");
                        break;

                    case ProcedureType.ProcedureFunction:
                        oResult.Append(procedure.ReturnType);
                        break;

                    case ProcedureType.ProcedureEvent:
                        oResult.Append("void");
                        break;
                }

                // name
                oResult.Append(" " + procedure.Name);
                // parameters
                if (procedure.ParameterList.Count <= 0)
                {
                    oResult.Append("()\r\n");
                }

                // start body
                oResult.Append(Indent4 + "{\r\n");

                foreach (string line in procedure.LineList)
                {
                    var temp = line.Trim();
                    if (temp.Length > 0)
                    {
                        oResult.Append(Indent6 + temp + ";\r\n");
                    }
                    else
                    {
                        oResult.Append("\r\n");
                    }
                }

                foreach (string line in procedure.BottomLineList)
                {
                    var temp = line.Trim();
                    if (temp.Length > 0)
                    {
                        oResult.Append(Indent6 + temp + ";\r\n");
                    }
                    else
                    {
                        oResult.Append("\r\n");
                    }
                }

                // end procedure
                oResult.Append(Indent4 + "}\r\n");
            }
        }

        public void ConvertPropertyCode(StringBuilder result)
        {
            // properties
            if (mTargetModule.PropertyList.Count > 0)
            {
                // new line
                result.Append("\r\n");
                //public string Comment
                //{
                //  get { return mComment; }
                //  set { mComment = value; }
                //}
                foreach (Property property in mTargetModule.PropertyList)
                {
                    // possible comment
                    result.Append(property.Comment + ";\r\n");
                    // string Result = null;
                    result.Append(Indent4 + property.Scope + " " + property.Type + " " + property.Name + " { get; set; }\r\n");

                    // lines
                    var atBottom = new List<string>();
                    foreach (string line in property.LineList)
                    {
                        var temp = line.Trim();
                        if (temp.Length > 0)
                        {
                            Tools.ConvertLineOfCode(temp, out string convertedLineOfCode, out var placeAtBottom);
                            if(convertedLineOfCode.IsNotEmpty())
                                convertedLineOfCode = Indent6 + convertedLineOfCode + ";";
                            result.AppendLine(convertedLineOfCode); 
                            if(placeAtBottom.IsNotEmpty())
                                atBottom.Add(placeAtBottom);
                        }
                        else
                        {
                            result.Append("\r\n");
                        }
                    }

                    foreach (string line in atBottom)
                        result.AppendLine(line);

                    result.Append(Indent4 + "}\r\n");
                }
            }
        }

        public void ConvertVariablesInCode(StringBuilder oResult)
        {
            oResult.Append("\r\n");

            foreach (Variable variable in mTargetModule.VariableList)
            {
                // string Result = null;
                oResult.Append(Indent4 + variable.Scope + " " + variable.Type + " " + variable.Name + ";\r\n");
            }
        }

        // generate result file
        // OutPath for pictures
        public string GetCode(string outPath)
        {
            string temp;

            var oResult = new StringBuilder();

            // convert source to target
            mTargetModule = new Module();
            Tools.ParseModule(mSourceModule, mTargetModule);

            // ********************************************************
            // common class
            // ********************************************************
            oResult.Append("using System;\r\n");

            // ********************************************************
            // only form class
            // ********************************************************
            if (mTargetModule.Type == "form")
            {
                oResult.Append("using System.Drawing;\r\n");
                oResult.Append("using System.Collections;\r\n");
                oResult.Append("using System.ComponentModel;\r\n");
                oResult.Append("using System.Windows.Forms;\r\n");
            }

            oResult.Append("\r\n");
            oResult.Append($"namespace {ProjectNamespace}\r\n");
            // start namepsace region
            oResult.Append("{\r\n");
            if (!string.IsNullOrEmpty(mSourceModule.Comment))
            {
                oResult.Append(Indent2 + "/// <summary>\r\n");
                oResult.Append(Indent2 + "///   " + mSourceModule.Comment + ".\r\n");
                oResult.Append(Indent2 + "/// </summary>\r\n");
            }

            switch (mTargetModule.Type)
            {
                case "form":
                    oResult.Append(Indent2 + "public class " + mSourceModule.Name + " : System.Windows.Forms.Form\r\n");
                    break;

                case "module":
                    oResult.Append(Indent2 + "class " + mSourceModule.Name + "\r\n");
                    // all procedures must be static
                    break;

                case "class":
                    oResult.Append(Indent2 + "public class " + mSourceModule.Name + "\r\n");
                    break;
            }
            // start class region
            oResult.Append(Indent2 + "{\r\n");

            // ********************************************************
            // only form class
            // ********************************************************

            if (mTargetModule.Type == "form")
            {
                ConvertFormCode(outPath, oResult);
            }

            // ********************************************************
            // enums
            // ********************************************************

            if (mTargetModule.EnumList.Count > 0)
            {
                ConvertEnumCode(oResult);
            }

            // ********************************************************
            //  variables for al module types
            // ********************************************************

            if (mTargetModule.VariableList.Count > 0)
            {
                ConvertVariablesInCode(oResult);
            }

            // ********************************************************
            // properties has only forms and classes
            // ********************************************************

            if ((mTargetModule.Type == "form") || (mTargetModule.Type == "class"))
            {
                ConvertPropertyCode(oResult);
            }

            // ********************************************************
            // procedures
            // ********************************************************

            if (mTargetModule.ProcedureList.Count > 0)
            {
                ConvertProcedureCode(oResult);
            }

            // end class
            oResult.Append(Indent2 + "}\r\n");
            // end namespace
            oResult.Append("}\r\n");

            // return result
            return oResult.ToString();
        }

        public void GetPropertyRow(StringBuilder result, string type, string name, ControlProperty controlProperty, string outPath)
        {
            // exception for images
            if (controlProperty.Name == "Icon" || controlProperty.Name == "Image" || controlProperty.Name == "BackgroundImage")
            {
                // generate resx file and write there image extracted from VB6 frx file
                var resourceName = string.Empty;

                switch (controlProperty.Name)
                {
                    case "BackgroundImage":
                        resourceName = "$this.BackgroundImage";
                        break;

                    case "Icon":
                        resourceName = "$this.Icon";
                        break;

                    case "Image":
                        resourceName = name + ".Image";
                        break;
                }

                if (!WriteImage(mSourceModule, resourceName, controlProperty.Value, outPath)) return;

                switch (controlProperty.Name)
                {
                    case "BackgroundImage":
                        result.Append(Indent6 + "this."
                                               + controlProperty.Name + " = ((System.Drawing.Bitmap)(resources.GetObject("
                                               + (char)34 + "$this.BackgroundImage" + (char)34 + ")));\r\n");
                        break;

                    case "Icon":
                        result.Append(Indent6 + "this."
                                               + controlProperty.Name + " = ((System.Drawing.Icon)(resources.GetObject("
                                               + (char)34 + "$this.Icon" + (char)34 + ")));\r\n");
                        break;

                    case "Image":
                        result.Append(Indent6 + "this." + name + "."
                                       + controlProperty.Name + " = ((System.Drawing.Bitmap)(resources.GetObject("
                                       + (char)34 + name + ".Image" + (char)34 + ")));\r\n");
                        break;
                }
            }
            else
            {
                // unsupported property
                if (!controlProperty.Valid)
                {
                    result.Append("//");
                }
                if (type == "form")
                {
                    // form properties
                    result.Append(Indent6 + "this."
                      + controlProperty.Name + " = " + controlProperty.Value + ";\r\n");
                }
                else
                {
                    // control properties
                    result.Append(Indent6 + "this." + name + "."
                      + controlProperty.Name + " = " + controlProperty.Value + ";\r\n");
                }
            }
        }

        public string GetWord(string line, ref int position)
        {
            string result = null;

            if (position < line.Length)
            {
                // seach for first space
                var end = line.IndexOf(" ", position);
                if (end > -1)
                {
                    result = line.Substring(position, end - position);
                    position = end;
                }
                else
                {
                    result = line.Substring(position);
                }
            }
            return result;
        }

        public bool ParseClass(StreamReader reader)
        {
            string line = null;
            var tempString = string.Empty;

            //VERSION 1.0 CLASS
            //BEGIN
            //  MultiUse = -1  'True
            //END
            //Attribute VB_Name = "CList"
            //Attribute VB_GlobalNameSpace = False
            //Attribute VB_Creatable = True
            //Attribute VB_PredeclaredId = False
            //Attribute VB_Exposed = True

            // Start from begin again
            reader.DiscardBufferedData();
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (reader.Peek() > -1)
            {
                var position = 0;
                // verify type of file based on first line
                // form, module, class
                line = reader.ReadLine();
                // next word - control type
                tempString = GetWord(line, ref position);
                if (tempString == "Attribute")
                {
                    position++;
                    tempString = GetWord(line, ref position);
                    switch (tempString)
                    {
                        case "VB_Name":
                            position++;
                            tempString = GetWord(line, ref position);
                            position++;
                            mSourceModule.Name = GetWord(line, ref position).Replace("\"", string.Empty);
                            break;

                        case "VB_Exposed":
                            return true;

                        case "VB_Description":
                            position++;
                            tempString = GetWord(line, ref position);
                            position++;
                            mSourceModule.Comment = GetWord(line, ref position).Replace("\"", string.Empty);
                            break;
                    }
                }
            }
            return false;
        }

        //Public Enum ENUM_BUG_LEVEL
        //  BUG_LEVEL_PROJECT = 1
        //  BUG_LEVEL_VERSION = 2
        //End Enum
        public void ParseEnumItem(EnumItem enumItem, string line)
        {
            var iPosition = 0;

            line = line.Trim();
            // first word is ame
            enumItem.Name = GetWord(line, ref iPosition);
            iPosition++;
            // next word =
            GetWord(line, ref iPosition);
            iPosition++;
            // optional
            enumItem.Value = GetWord(line, ref iPosition);
        }

        public bool ParseForm(StreamReader oReader)
        {
            var bProcess = false;
            var bFinish = false;
            var bEnd = true;
            string sOwner = null;
            var iComment = 0;
            string sType = null;
            var iPosition = 0;
            var iLevel = 0;
            Control control = null;
            ControlProperty oNestedProperty = null;
            var bNestedProperty = false;

            // parse only visual part of form
            while (!bFinish) // ( ( bFinish || (oReader.Peek() > -1)) )
            {
                var sLine = oReader.ReadLine();
                sLine = sLine.Trim();
                iPosition = 0;
                // get first word in line
                var sTemp = GetWord(sLine, ref iPosition);
                string sName = null;
                switch (sTemp)
                {
                    case "Begin":
                        bProcess = true;
                        // new level
                        iLevel++;
                        // next word - control type
                        iPosition++;
                        sType = GetWord(sLine, ref iPosition);
                        // next word - control name
                        iPosition++;
                        sName = GetWord(sLine, ref iPosition);
                        // detected missing end -> it indicate that next control is container
                        if (!bEnd)
                        {
                            // add container control to colection
                            if (!(control == null))
                            {
                                control.Container = true;
                                mSourceModule.ControlAdd(control);
                            }
                            // save name of previous control as owner for current and next controls
                            moOwnerStock.Add(sOwner);
                        }
                        bEnd = false;

                        switch (sType)
                        {
                            case "Form":            // VERSION 2.00 - VB3
                            case "VB.Form":
                            case "VB.MDIForm":
                                mSourceModule.Name = sName;
                                // first owner
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                break;

                            default:
                                // new control
                                control = new Control();
                                control.Name = sName;
                                control.Type = sType;
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                // set current container name
                                control.Owner = (string)moOwnerStock[moOwnerStock.Count - 1];
                                break;
                        }
                        break;

                    case "End":
                        // double end - we leaving some container
                        if (bEnd)
                        {
                            // remove last item from stock
                            moOwnerStock.Remove((string)moOwnerStock[moOwnerStock.Count - 1]);
                        }
                        else
                        {
                            // level 1 is form and all higher levels are controls
                            if (iLevel > 1)
                            {
                                // add control to colection
                                mSourceModule.ControlAdd(control);
                            }
                        }
                        // form or control end detected
                        bEnd = true;
                        // back to previous level
                        iLevel--;

                        break;

                    case "Object":
                        // used controls in form
                        break;

                    case "BeginProperty":
                        bNestedProperty = true;

                        oNestedProperty = new ControlProperty();
                        // next word - nested property name
                        iPosition++;
                        sName = GetWord(sLine, ref iPosition);
                        oNestedProperty.Name = sName;
                        //            Debug.WriteLine(sName);
                        break;

                    case "EndProperty":
                        bNestedProperty = false;
                        // add property to control or form
                        if (iLevel == 1)
                        {
                            // add property to form
                            mSourceModule.FormPropertyAdd(oNestedProperty);
                        }
                        else
                        {
                            // to controls
                            control.PropertyAdd(oNestedProperty);
                        }
                        break;

                    default:
                        // parse property
                        var property = new ControlProperty();

                        var iTemp = sLine.IndexOf("=");
                        if (iTemp > -1)
                        {
                            property.Name = sLine.Substring(0, iTemp - 1).Trim();
                            iComment = sLine.IndexOf("'", iTemp);
                            if (iComment > -1)
                            {
                                property.Value = sLine.Substring(iTemp + 1, iComment - iTemp - 1).Trim();
                                property.Comment = sLine.Substring(iComment + 1, sLine.Length - iComment - 1).Trim();
                            }
                            else
                            {
                                property.Value = sLine.Substring(iTemp + 1, sLine.Length - iTemp - 1).Trim();
                            }

                            if (bNestedProperty)
                            {
                                oNestedProperty.PropertyList.Add(property);
                            }
                            else
                            {
                                // depend by level insert property to form or control
                                if (iLevel > 1)
                                {
                                    // add property to control
                                    control.PropertyAdd(property);
                                }
                                else
                                {
                                    // add property to form
                                    mSourceModule.FormPropertyAdd(property);
                                }
                            }
                        }
                        break;
                }

                if ((iLevel == 0) && bProcess)
                {
                    // visual part of form is finish
                    bFinish = true;
                }
            }
            return true;
        }

        public bool ParseModule(StreamReader reader)
        {
            // Start from begin again
            reader.DiscardBufferedData();
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            // search for module name
            while (reader.Peek() > -1)
            {
                var line = reader.ReadLine();
                var position = line.IndexOf('"');
                mSourceModule.Name = line.Substring(position + 1, line.Length - position - 2);
                return true;
            }
            return false;
        }

        public void ParseParameters(List<Parameter> parameterList, string line)
        {
            var bFinish = false;
            var position = 0;
            var start = 0;
            var status = false;
            var tempString = string.Empty;

            // parameters delimited by comma
            while (!bFinish)
            {
                var parameter = new Parameter();

                // next word - control type
                tempString = GetWord(line, ref position);
                switch (tempString)
                {
                    case "Optional":

                        break;

                    case "ByVal":
                    case "ByRef":
                        parameter.Pass = tempString;
                        break;
                    // missing is byref
                    default:
                        parameter.Pass = "ByRef";
                        // variable name
                        parameter.Name = tempString;
                        status = true;
                        break;
                }

                // variable name
                if (!status)
                {
                    position++;
                    tempString = GetWord(line, ref position);
                    parameter.Name = tempString;
                }
                // As
                position++;
                tempString = GetWord(line, ref position);
                // parameter type
                position++;
                tempString = GetWord(line, ref position);
                parameter.Type = tempString;

                parameterList.Add(parameter);

                // next parameter
                position++;
                start = position;

                if (start <= line.Length)
                    position = line.IndexOf(",", start);
                else
                    position = -1;

                if (position == -1)
                {
                    // end
                    bFinish = true;
                }
            }
        }

        // ByVal lValue As Long, ByVal sValue As string
        public void ParseProcedureName(Procedure procedure, string line)
        {
            var tempString = string.Empty;
            var iPosition = 0;
            var start = 0;
            var status = false;

            //public Sub cmdOk_Click()
            //public void cmdShow_Click(object sender, System.EventArgs e)

            //public Sub Form_Load()
            //public void frmConvert_Load(object sender, System.EventArgs e)

            //Public Function Rozbor_DefaultFields(ByVal MKf As String) As String
            //public static bool ParseProcedures( Module SourceModule, Module TargetModule )

            tempString = GetWord(line, ref iPosition);
            switch (tempString)
            {
                case "Private":
                    procedure.Scope = "private";
                    status = true;
                    break;

                case "Public":
                    procedure.Scope = "public";
                    status = true;
                    break;

                default:
                    procedure.Scope = "private";
                    status = true;
                    break;
            }

            if (status)
            {
                // property
                iPosition++;
                tempString = GetWord(line, ref iPosition);
            }

            // procedure type
            switch (tempString)
            {
                case "Sub":
                    procedure.Type = ProcedureType.ProcedureSub;
                    break;

                case "Function":
                    procedure.Type = ProcedureType.ProcedureFunction;
                    break;

                case "Event":
                    procedure.Type = ProcedureType.ProcedureEvent;
                    break;
            }

            // next is name
            iPosition++;
            start = iPosition;
            iPosition = line.IndexOf("(", start);
            procedure.Name = line.Substring(start, iPosition - start);

            // next possible parameters
            iPosition++;
            start = iPosition;
            iPosition = line.IndexOf(")", start);

            if ((iPosition - start) > 0)
            {
                tempString = line.Substring(start, iPosition - start);
                var parameterList = new List<Parameter>();
                // process parameters
                ParseParameters(parameterList, tempString);
                procedure.ParameterList = parameterList;
            }

            // and return type of function
            if (procedure.Type == ProcedureType.ProcedureFunction)
            {
                // as
                iPosition++;
                tempString = GetWord(line, ref iPosition);
                // function return type
                iPosition++;
                procedure.ReturnType = GetWord(line, ref iPosition);
            }
        }

        public bool ParseProcedures(StreamReader reader)
        {
            string line = null;
            string tempString = null;
            string sComments = null;
            string sScope = null;

            var iPosition = 0;
            //bool bProcess = false;

            var bEnum = false;
            var bVariable = false;
            var bProperty = false;
            var bProcedure = false;
            var bEnd = false;

            Variable variable = null;
            Property property = null;
            Procedure procedure = null;
            Enum enumItems = null;
            EnumItem enumItem = null;

            while (reader.Peek() > -1)
            {
                line = reader.ReadLine();
                iPosition = 0;

                if (line.IsNotEmpty())
                {
                    // check if next line is same command, join it together ?
                    while (line.Substring(line.Length - 1, 1) == "_")
                    {
                        line += reader.ReadLine();
                    }
                }

                // get first word in line
                tempString = GetWord(line, ref iPosition);
                if (int.TryParse(tempString, out var tempResult))
                {
                    line = line.Substring(tempResult.ToString().Length).Trim();
                    iPosition = 0;
                    tempString = GetWord(line, ref iPosition);
                }

                switch (tempString)
                {
                    // ignore this section

                    //Attribute VB_Name = "frmAttachement"
                    // ...
                    //Option Explicit
                    case "Attribute":
                    case "Option":
                        break;

                    // comments
                    case "'":
                        sComments = sComments + line + "\r\n";
                        break;

                    // next can be declaration of variables

                    //public mlParentID As Long
                    //public mlOwnerType As ENUM_FORM_TYPE
                    //public moAttachement As Attachement

                    case "Public":
                    case "Private":
                        // save it for later use
                        sScope = tempString.ToLower();
                        // read next word
                        // next word - control type
                        iPosition++;
                        tempString = GetWord(line, ref iPosition);

                        switch (tempString)
                        {
                            // functions or procedures
                            case "Sub":
                            case "Function":

                                procedure = new Procedure();
                                procedure.Comment = sComments;
                                sComments = string.Empty;
                                ParseProcedureName(procedure, line);

                                bProcedure = true;
                                break;

                            case "Enum":
                                enumItems = new Enum();
                                enumItems.Scope = sScope;
                                // next word is enum name
                                iPosition++;
                                enumItems.Name = GetWord(line, ref iPosition);
                                bEnum = true;
                                break;

                            case "Property":
                                property = new Property();
                                property.Comment = sComments ?? string.Empty;
                                sComments = string.Empty;
                                ParsePropertyName(property, line);
                                bProperty = true;

                                break;

                            default:
                                // variable declaration
                                variable = new Variable();
                                ParseVariableDeclaration(variable, line);
                                bVariable = true;
                                break;
                        }

                        break;

                    case "Dim":
                        // variable declaration
                        variable = new Variable();
                        variable.Comment = sComments;
                        sComments = string.Empty;
                        ParseVariableDeclaration(variable, line);
                        bVariable = true;
                        break;

                    // functions or procedures
                    case "Sub":
                    case "Function":
                        break;

                    case "End":
                        bEnd = true;
                        break;

                    default:
                        if (bEnum)
                        {
                            // first word is name, second =, thirt value if is preset
                            enumItem = new EnumItem();
                            enumItem.Comment = sComments;
                            sComments = string.Empty;
                            ParseEnumItem(enumItem, line);
                            // add item
                            enumItems.ItemList.Add(enumItem);
                        }
                        if (bProperty)
                        {
                            // add line of property
                            property.LineList.Add(line);
                        }
                        if (bProcedure)
                        {
                            procedure.LineList.Add(line);
                        }
                        break;

                        // events
                        //public Sub cmdCancel_Click()
                        //  mbEdit = False
                        //  If mbNew Then
                        //    Unload Me
                        //  Else
                        //    ShowCurRec
                        //    SetControls False
                        //  End If
                        //End Sub
                        //
                        //public Sub cmdClose_Click()
                        //  Unload Me
                        //End Sub
                }

                // if something end
                if (bEnd)
                {
                    //
                    if (bEnum)
                    {
                        mSourceModule.EnumList.Add(enumItems);
                        bEnum = false;
                    }
                    if (bProperty)
                    {
                        mSourceModule.PropertyAdd(property);
                        bProperty = false;
                    }
                    if (bProcedure)
                    {
                        mSourceModule.ProcedureAdd(procedure);
                        bProcedure = false;
                    }
                    bEnd = false;
                }
                else
                {
                    if (bVariable)
                    {
                        mSourceModule.VariableAdd(variable);
                    }
                }

                bVariable = false;
            }

            return true;
        }

        //public mlID As Long

        public void ParsePropertyName(Property property, string line)
        {
            var tempString = string.Empty;
            var iPosition = 0;
            var start = 0;
            var status = false;

            // next word - control type
            tempString = GetWord(line, ref iPosition);
            switch (tempString)
            {
                case "Private":
                    property.Scope = "private";
                    break;

                case "Public":
                    property.Scope = "public";
                    break;

                default:
                    property.Scope = "private";
                    status = true;
                    break;
            }

            if (!status)
            {
                // property
                iPosition++;
                tempString = GetWord(line, ref iPosition);
            }

            // direction Let,Get, Set
            iPosition++;
            tempString = GetWord(line, ref iPosition);
            property.Direction = tempString;

            //Public Property Let ParentID(ByVal lValue As Long)

            // name
            start = iPosition;
            iPosition = line.IndexOf("(", start + 1);
            property.Name = line.Substring(start, iPosition - start);

            // + possible parameters
            iPosition++;
            start = iPosition;
            iPosition = line.IndexOf(")", start);

            if ((iPosition - start) > 0)
            {
                tempString = line.Substring(start, iPosition - start);
                // process parameters
                ParseParameters(property.ParameterList, tempString);
            }

            // As
            iPosition++;
            iPosition++;
            tempString = GetWord(line, ref iPosition);

            // type
            iPosition++;
            tempString = GetWord(line, ref iPosition);
            property.Type = tempString;
        }

        public void ParseVariableDeclaration(Variable variable, string line)
        {
            var tempString = string.Empty;
            var iPosition = 0;
            var status = false;

            // next word - control type
            tempString = GetWord(line, ref iPosition);
            switch (tempString)
            {
                case "Dim":
                case "Private":
                    variable.Scope = "private";
                    break;

                case "Public":
                    variable.Scope = "public";
                    break;

                default:
                    variable.Scope = "private";
                    // variable name
                    variable.Name = tempString;
                    status = true;
                    break;
            }

            // variable name
            if (!status)
            {
                iPosition++;
                tempString = GetWord(line, ref iPosition);
                variable.Name = tempString;
            }
            // As
            iPosition++;
            tempString = GetWord(line, ref iPosition);
            // variable type
            iPosition++;
            tempString = GetWord(line, ref iPosition);
            variable.Type = tempString == "String" ? "string" : tempString;
        }

        public bool WriteImage(Module sourceModule, string resourceName, string value, string outPath)
        {
            var temp = string.Empty;
            var offset = 0;
            var frxFile = string.Empty;
            var sResxName = string.Empty;
            var position = 0;

            position = value.IndexOf(":", 0);
            // "Form1.frx":0000;
            // old vb3 code has name without ""
            // CONTROL.FRX:0000

            if (sourceModule.Version == "5.00")
            {
                frxFile = value.Substring(1, position - 2);
            }
            else
            {
                frxFile = value.Substring(0, position);
            }

            temp = value.Substring(position + 1, value.Length - position - 1);
            offset = Convert.ToInt32("0x" + temp, 16);
            // exist file ?

            // get image
            byte[] imageString;

            Tools.GetFrxImage(Path.GetDirectoryName(sourceModule.FileName) + @"\" + frxFile, offset, out imageString);

            if ((imageString.GetLength(0) - 8) > 0)
            {
                if (File.Exists(outPath + resourceName))
                {
                    File.Delete(outPath + resourceName);
                }
                FileStream stream = stream = new FileStream(outPath + resourceName, FileMode.CreateNew, FileAccess.Write);
                var writer = new BinaryWriter(stream);
                writer.Write(imageString, 8, imageString.GetLength(0) - 8);
                stream.Close();
                writer.Close();

                // write
                mTargetModule.ImageList.Add(resourceName);
                return true;
            }
            else
            {
                return false;
            }
            // save it to resx
            //      Debug.WriteLine(ModuleName + ", " + ResourceName + ", " + Temp + ", " + oImage.Width.ToString() );
            //
            //      ResXResourceWriter oWriter;
            //
            //      sResxName = @"C:\temp\test\" + ModuleName + ".resx";
            //      oWriter = new ResXResourceWriter(sResxName);
            //     // oImage = Image.FromFile(myRow[sField].ToString());
            //      oWriter.AddResource(ResourceName, oImage);
            //      oWriter.Generate();
            //      oWriter.Close();

            //      MemoryStream s = new MemoryStream();
            //      oImage.Save(s);
            //
            //      byte[] b = s.ToArray();
            //      String strContents = System.Security.Cryptography.EncodeAsBase64.EncodeBuffer(b);
            //      myXmlTextNode.Value = strContents;

            /* Add Convert it back using.... */
            //
            //byte[]  b =
            //System.Security.Cryptography.DecodeBase64.DecodeBuffer(myXmlTextNode.Value);

            //      ResourceWriter oWriter;
            //
            //      sResxName = @"c:\temp\" + ModuleName + ".resx";
            //      oWriter = new ResourceWriter(sResxName);
            //     // oImage = Image.FromFile(myRow[sField].ToString());
            //      oWriter.AddResource(ResourceName, oImage);
            //      oWriter.Close();
        }

        // properties Let, Get, Set
        //Public Property Let ParentID(ByVal lValue As Long)
        //  mlParentID = lValue
        //End Property
        //
        //Public Property Get FormType() As ENUM_FORM_TYPE
        //  FormType = FORM_ATTACHEMENT
        //End Property
        public void WriteResX(List<string> mImageList, string outPath, string moduleName)
        {
            string sResxName;

            if (mImageList.Count > 0)
            {
                // resx name
                sResxName = outPath + moduleName + ".resx";
                // open file
                var rsxw = new ResXResourceWriter(sResxName);

                foreach (string resourceName in mImageList)
                {
                    try
                    {
                        var img = Image.FromFile(outPath + resourceName);
                        rsxw.AddResource(resourceName, img);
                        img.Dispose();
                    }
                    catch
                    {
                    }
                }
                // rsxw.Generate();
                rsxw.Close();

                foreach (string resourceName in mImageList)
                {
                    File.Delete(outPath + resourceName);
                }
            }
        }
    }
}