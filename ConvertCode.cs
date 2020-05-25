using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        public readonly List<string> OwnerStock;
        public VbFileType FileType;
        public Module SourceModule;
        public Module TargetModule;
        public string ActionResult;

        public string Code;

        public string ProjectNamespace = "MetX.SliceAndDice";

        public ConvertCode()
        {
            OwnerStock = new List<string>();
        }

        public void ConvertEnumCode(StringBuilder oResult)
        {
            oResult.AppendLine();
            foreach (var enumItems in TargetModule.EnumList)
            {
                // public enum VB_FILE_TYPE
                oResult.Append(Indent4 + enumItems.Scope + " enum " + enumItems.Name + "\r\n");
                oResult.Append(Indent4 + "{\r\n");

                foreach (var enumItem in enumItems.ItemList)
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

        public void ConvertFormCode(string outPath, StringBuilder oResult)
        {
            // list of controls
            foreach (var control in TargetModule.ControlList)
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
            oResult.AppendLine();
            oResult.Append(Indent4 + "public " + SourceModule.Name + "()\r\n");
            oResult.Append(Indent4 + "{\r\n");
            oResult.Append(Indent6 + "// Required for Windows Form Designer support\r\n");
            oResult.Append(Indent6 + "InitializeComponent();\r\n");
            oResult.AppendLine();
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
            if (TargetModule.ImagesUsed)
            {
                // System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
                oResult.Append(Indent6 + "System.Resources.ResourceManager resources = " +
                               "new System.Resources.ResourceManager(typeof(" + TargetModule.Name + "));\r\n");
            }

            foreach (var control in TargetModule.ControlList)
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
            foreach (var control in TargetModule.ControlList)
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
            foreach (var control in TargetModule.ControlList)
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
                foreach (var property in control.PropertyList)
                {
                    GetPropertyRow(oResult, control.Type, control.Name, property, outPath);
                }

                // if control is container for other controls
                var temp = string.Empty;
                foreach (var oControl1 in TargetModule.ControlList)
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
            oResult.Append(Indent6 + "// " + SourceModule.Name + "\r\n");
            oResult.Append(Indent6 + "//\r\n");
            oResult.Append(Indent6 + "this.Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
            oResult.Append(Indent6 + "{\r\n");

            // add control range to form
            foreach (var control in TargetModule.ControlList)
            {
                if (!control.Valid)
                {
                    oResult.Append("//");
                }

                // all controls ownered by main form
                if ((control.Owner == SourceModule.Name) && (!control.InvisibleAtRuntime))
                {
                    oResult.Append(Indent6 + "      this." + control.Name + ",\r\n");
                }
            }

            // remove last comma, keep CRLF
            oResult.Remove(oResult.Length - 3, 1);
            // close addrange part
            oResult.Append(Indent6 + "});\r\n");

            // form name
            oResult.Append(Indent6 + "this.Name = " + (char)34 + TargetModule.Name + (char)34 + ";\r\n");
            // exception for menu
            // this.Menu = this.mainMenu1;
            if (TargetModule.MenuUsed)
            {
                foreach (var control in TargetModule.ControlList)
                {
                    if (control.Type == "MainMenu")
                    {
                        oResult.Append(Indent6 + "      this.Menu = " + control.Name + ";\r\n");
                    }
                }
            }

            // form properties
            foreach (var property in TargetModule.FormPropertyList)
            {
                if (!property.Valid)
                {
                    oResult.Append("//");
                }

                GetPropertyRow(oResult, TargetModule.Type, "", property, outPath);
            }

            // resume layout for each container
            foreach (var control in TargetModule.ControlList)
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

        public void ConvertProcedureCode(StringBuilder oResult)
        {
            oResult.AppendLine();
            foreach (var procedure in TargetModule.ProcedureList)
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

                foreach (var line in procedure.LineList)
                {
                    var temp = line.Trim();
                    if (temp.Length > 0)
                    {
                        oResult.Append(Indent6 + temp + ";\r\n");
                    }
                    else
                    {
                        oResult.AppendLine();
                    }
                }

                foreach (var line in procedure.BottomLineList)
                {
                    var temp = line.Trim();
                    if (temp.Length > 0)
                    {
                        oResult.Append(Indent6 + temp + ";\r\n");
                    }
                    else
                    {
                        oResult.AppendLine();
                    }
                }

                // end procedure
                oResult.Append(Indent4 + "}\r\n");
            }
        }

        public void ConvertPropertyCode(StringBuilder result)
        {
            // properties
            if (TargetModule.PropertyList.Count > 0)
            {
                // new line
                result.AppendLine();
                //public string Comment
                //{
                //  get { return mComment; }
                //  set { mComment = value; }
                //}
                foreach (var property in TargetModule.PropertyList)
                {
                    // possible comment
                    result.Append(property.Comment + ";\r\n");
                    // string Result = null;
                    result.Append(Indent4 + property.Scope + " " + property.Type + " " + property.Name + ";\r\n");

                    // lines
                    var atBottom = new List<string>();
                    foreach (var line in property.LineList)
                    {
                        var temp = line.Trim();
                        if (temp.Length > 0)
                        {
                            Tools.ConvertLineOfCode(temp, out var convertedLineOfCode, out var placeAtBottom);
                            if (convertedLineOfCode.IsNotEmpty())
                                convertedLineOfCode = Indent6 + convertedLineOfCode + ";";
                            result.AppendLine(convertedLineOfCode);
                            if (placeAtBottom.IsNotEmpty())
                                atBottom.Add(placeAtBottom);
                        }
                        else
                        {
                            result.AppendLine();
                        }
                    }

                    foreach (var line in atBottom)
                        result.AppendLine(line);

                    result.Append(Indent4 + "}\r\n");
                }
            }
        }

        public void ConvertVariablesInCode(StringBuilder oResult)
        {
            oResult.AppendLine();

            foreach (var variable in TargetModule.VariableList)
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

            var result = new StringBuilder();

            // convert source to target
            TargetModule = new Module();
            Tools.ParseModule(SourceModule, TargetModule);

            // ********************************************************
            // common class
            // ********************************************************
            result.Append("using System;\r\n");

            // ********************************************************
            // only form class
            // ********************************************************
            if (TargetModule.Type == "form")
            {
                result.Append("using System.Drawing;\r\n");
                result.Append("using System.Collections;\r\n");
                result.Append("using System.ComponentModel;\r\n");
                result.Append("using System.Windows.Forms;\r\n");
            }

            result.AppendLine();
            result.AppendLine($"namespace {ProjectNamespace}");
            // start namepsace region
            result.Append("{\r\n");
            if (!string.IsNullOrEmpty(SourceModule.Comment))
            {
                result.Append(Indent2 + "/// <summary>\r\n");
                result.Append(Indent2 + "///   " + SourceModule.Comment + ".\r\n");
                result.Append(Indent2 + "/// </summary>\r\n");
            }

            switch (TargetModule.Type)
            {
                case "form":
                    result.AppendLine(Indent2 + "public class " + SourceModule.Name + " : System.Windows.Forms.Form");
                    break;

                case "module":
                    result.AppendLine(Indent2 + "class " + SourceModule.Name);
                    // all procedures must be static
                    break;

                case "class":
                    result.AppendLine(Indent2 + "public class " + SourceModule.Name);
                    break;
            }
            // start class region
            result.Append(Indent2 + "{\r\n");

            // ********************************************************
            // only form class
            // ********************************************************

            if (TargetModule.Type == "form")
            {
                ConvertFormCode(outPath, result);
            }

            // ********************************************************
            // enums
            // ********************************************************

            if (TargetModule.EnumList.Count > 0)
            {
                ConvertEnumCode(result);
            }

            // ********************************************************
            //  variables for al module types
            // ********************************************************

            if (TargetModule.VariableList.Count > 0)
            {
                ConvertVariablesInCode(result);
            }

            // ********************************************************
            // properties has only forms and classes
            // ********************************************************

            if ((TargetModule.Type == "form") || (TargetModule.Type == "class"))
            {
                ConvertPropertyCode(result);
            }

            // ********************************************************
            // procedures
            // ********************************************************

            if (TargetModule.ProcedureList.Count > 0)
            {
                ConvertProcedureCode(result);
            }

            // end class
            result.Append(Indent2 + "}\r\n");
            // end namespace
            result.Append("}\r\n");

            // return result
            return result.ToString();
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

                if (!WriteImage(SourceModule, resourceName, controlProperty.Value, outPath)) return;

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
                            SourceModule.Name = GetWord(line, ref position).Replace("\"", string.Empty);
                            break;

                        case "VB_Exposed":
                            return true;

                        case "VB_Description":
                            position++;
                            tempString = GetWord(line, ref position);
                            position++;
                            SourceModule.Comment = GetWord(line, ref position).Replace("\"", string.Empty);
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

        public bool ParseFile(string filename, string outputPath)
        {
            var version = string.Empty;
            bool result;

            // try recognize source code type depend by file extension
            var extension = filename.Substring(filename.Length - 3, 3);
            switch (extension.ToUpper())
            {
                case "FRM":
                    FileType = VbFileType.VbFileForm;
                    break;

                case "BAS":
                    FileType = VbFileType.VbFileModule;
                    break;

                case "CLS":
                    FileType = VbFileType.VbFileClass;
                    break;

                default:
                    FileType = VbFileType.VbFileUnknown;
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
                            FileType = VbFileType.VbFileModule;
                            break;

                        // form or class first line
                        // 'VERSION 5.00' or 'VERSION 1.0 CLASS'
                        case "VERSION":
                            position++;
                            version = GetWord(line, ref position);

                            FileType = line.Contains(ClassFirstLine)
                                ? VbFileType.VbFileClass
                                : VbFileType.VbFileForm;
                            break;

                        default:
                            FileType = VbFileType.VbFileUnknown;
                            break;
                    }

                    // if file is still unknown
                    if (FileType == VbFileType.VbFileUnknown)
                    {
                        ActionResult = "Unknown file type";
                        return false;
                    }

                    SourceModule = new Module
                    {
                        Version = version ?? "1.0",
                        FileName = filename
                    };

                    // now parse specifics of each type
                    switch (extension.ToUpper())
                    {
                        case "FRM":
                            SourceModule.Type = "form";
                            result = ParseForm(reader);
                            break;

                        case "BAS":
                            SourceModule.Type = "module";
                            result = ParseModule(reader);
                            break;

                        case "CLS":
                            SourceModule.Type = "class";
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
            var outFileName = outputPath + TargetModule.FileName;
            File.WriteAllText(outFileName, Code);

            // generate resx file if source form contain any images
            if ((TargetModule.ImagesUsed))
            {
                WriteResX(TargetModule.ImageList, outputPath, TargetModule.Name);
            }

            return result;
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
                                SourceModule.ControlAdd(control);
                            }
                            // save name of previous control as owner for current and next controls
                            OwnerStock.Add(sOwner);
                        }
                        bEnd = false;

                        switch (sType)
                        {
                            case "Form":            // VERSION 2.00 - VB3
                            case "VB.Form":
                            case "VB.MDIForm":
                                SourceModule.Name = sName;
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
                                control.Owner = (string)OwnerStock[OwnerStock.Count - 1];
                                break;
                        }
                        break;

                    case "End":
                        // double end - we leaving some container
                        if (bEnd)
                        {
                            // remove last item from stock
                            OwnerStock.Remove((string)OwnerStock[OwnerStock.Count - 1]);
                        }
                        else
                        {
                            // level 1 is form and all higher levels are controls
                            if (iLevel > 1)
                            {
                                // add control to colection
                                SourceModule.ControlAdd(control);
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
                            SourceModule.FormPropertyAdd(oNestedProperty);
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
                                    SourceModule.FormPropertyAdd(property);
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
                SourceModule.Name = line.Substring(position + 1, line.Length - position - 2);
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
            var word = string.Empty;
            // parameters delimited by comma
            var parts = line.Split(',').ToList().Select(p => p.Trim()).ToList();

            for (var index = 0; index < parts.Count; index++)
            {
                var part = parts[index];
                var parameter = new Parameter();
                
                parameter.Optional = part.StartsWith("Optional");
                part = part.Replace("Optional", "").Trim();
                
                parameter.Pass = part.StartsWith("ByVal") ? "ByVal" : "ByRef";
                part = part.Replace("ByRef ", "");
                part = part.Replace("ByVal ", "").Trim();
                parameter.Name = part.TokenAt(1, " As ").Trim();
                var type = part.TokenAt(2, " As ").Trim().TokenAt(1).Trim();
                parameter.Type = type;
                var left = part.TokenAt(2, " As ").Replace(type, "");
                if (left.IsNotEmpty())
                {
                    parameter.DefaultValue = left.Replace("= ", "");
                }

                parameterList.Add(parameter);

                // next parameter
                position++;
            }
        }

        // ByVal lValue As Long, ByVal sValue As string
        public void ParseProcedureName(Procedure procedure, string line)
        {
            var word = string.Empty;
            var position = 0;
            var start = 0;
            var status = false;

            //public Sub cmdOk_Click()
            //public void cmdShow_Click(object sender, System.EventArgs e)

            //public Sub Form_Load()
            //public void frmConvert_Load(object sender, System.EventArgs e)

            //Public Function Rozbor_DefaultFields(ByVal MKf As String) As String
            //public static bool ParseProcedures( Module SourceModule, Module TargetModule )

            word = GetWord(line, ref position);
            switch (word)
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
                position++;
                word = GetWord(line, ref position);
            }

            // procedure type
            switch (word)
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
            position++;
            start = position;
            position = line.IndexOf("(", start);
            procedure.Name = line.Substring(start, position - start);

            // next possible parameters
            position++;
            start = position;
            position = line.IndexOf(")", start);

            if ((position - start) > 0)
            {
                word = line.Substring(start, position - start);
                var parameterList = new List<Parameter>();
                // process parameters
                ParseParameters(parameterList, word);
                procedure.ParameterList = parameterList;
            }

            // and return type of function
            if (procedure.Type == ProcedureType.ProcedureFunction)
            {
                // as
                position++;
                word = GetWord(line, ref position);
                // function return type
                position++;
                procedure.ReturnType = GetWord(line, ref position);
            }
        }

        public bool ParseProcedures(StreamReader reader)
        {
            string line = null;
            string word = null;
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
                word = GetWord(line, ref iPosition);
                if (int.TryParse(word, out var tempResult))
                {
                    line = line.TokensAfter().Trim();
                    //line = line.Substring(tempResult.ToString().Length).Trim();
                    iPosition = 0;
                    word = GetWord(line, ref iPosition);
                }

                switch (word)
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
                        sScope = word.ToLower();
                        // read next word
                        // next word - control type
                        iPosition++;
                        word = GetWord(line, ref iPosition);

                        switch (word)
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
                        SourceModule.EnumList.Add(enumItems);
                        bEnum = false;
                    }
                    if (bProperty)
                    {
                        SourceModule.PropertyAdd(property);
                        bProperty = false;
                    }
                    if (bProcedure)
                    {
                        SourceModule.ProcedureAdd(procedure);
                        bProcedure = false;
                    }
                    bEnd = false;
                }
                else
                {
                    if (bVariable)
                    {
                        SourceModule.VariableAdd(variable);
                    }
                }

                bVariable = false;
            }

            return true;
        }

        //public mlID As Long

        public void ParsePropertyName(Property property, string line)
        {
            var word = string.Empty;
            var iPosition = 0;
            var start = 0;
            var status = false;

            // next word - control type
            word = GetWord(line, ref iPosition);
            switch (word)
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
                word = GetWord(line, ref iPosition);
            }

            // direction Let,Get, Set
            iPosition++;
            word = GetWord(line, ref iPosition);
            property.Direction = word;

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
                word = line.Substring(start, iPosition - start);
                // process parameters
                ParseParameters(property.ParameterList, word);
            }

            // As
            iPosition++;
            iPosition++;
            word = GetWord(line, ref iPosition);

            // type
            iPosition++;
            word = GetWord(line, ref iPosition);
            property.Type = word;
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
                TargetModule.ImageList.Add(resourceName);
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

                foreach (var resourceName in mImageList)
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

                foreach (var resourceName in mImageList)
                {
                    File.Delete(outPath + resourceName);
                }
            }
        }
    }
}