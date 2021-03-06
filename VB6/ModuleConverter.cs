using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using MetX.Library;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    ///     Summary description for Convert.
    /// </summary>
    public class ModuleConverter
    {
        public const string ClassFirstLine = "1.0 CLASS";
        public const string ModuleFirstLine = "ATTRIBUTE";
        public static readonly string Indent2 = Tools.Indent(2);
        public static readonly string Indent3 = Tools.Indent(3);
        public readonly List<string> OwnerStock = new List<string>();
        public string ActionResult;
        public VbFileType FileType;

        public ICodeLine FirstParent;
        public string ProjectNamespace = "MetX.SliceAndDice";
        public Module SourceModule;
        public Module TargetModule;

        public ModuleConverter(ICodeLine firstParent)
        {
            FirstParent = firstParent;
        }

        public void SetIndentsRecursively()
        {
            TargetModule.ResetIndent(0);

            foreach (var @enum in TargetModule.EnumList)
            {
            }

            foreach (var control in TargetModule.ControlList) control.ResetIndent(1);

            foreach (var procedure in TargetModule.ProcedureList) procedure.ResetIndent(1);

            foreach (var property in TargetModule.PropertyList) property.ResetIndent(1);

            foreach (var variable in TargetModule.VariableList) variable.ResetIndent(1);

            foreach (var formProperty in TargetModule.FormPropertyList) formProperty.ResetIndent(1);
        }


        public string ConvertEnumCode(int indentLevel)
        {
            var result = new StringBuilder();
            result.AppendLine();
            foreach (var enumItems in TargetModule.EnumList)
            {
                // public enum VB_FILE_TYPE
                result.Append(Indent2 + enumItems.Scope + " enum " + enumItems.Name + "\r\n");
                result.Append(Indent2 + "{\r\n");

                foreach (var enumItem in enumItems.ItemList)
                {
                    // name
                    result.Append(Indent3 + enumItem.Name);

                    if (enumItem.Value != string.Empty) result.Append(" = " + enumItem.Value);

                    // enum items delimiter
                    result.Append(",\r\n");
                }

                // remove last comma, keep CRLF
                result.Remove(result.Length - 3, 1);
                // end enum
                result.Append(Indent2 + "};\r\n");
            }

            var code = result.ToString();
            return code;
        }

        public bool ConvertFile(string filename, string outputPath)
        {
            var code = GenerateCode(filename, outputPath);
            if (code.IsEmpty())
                return false;

            // save result
            var outFileName = Path.Combine(outputPath, TargetModule.FileName);
            File.WriteAllText(outFileName, code);

            // generate resx file if source form contain any images
            if (TargetModule.ImagesUsed)
                WriteResX(TargetModule.ImageList, outputPath, TargetModule.Name);

            return code.IsNotEmpty();
        }

        public string GenerateCodeFragment(string vb6Code)
        {
            if (vb6Code.IsEmpty())
                return "";

            SourceModule = new Module(FirstParent)
            {
                Version = "1.0",
                FileName = "codeFragment",
                Type = "codeFragment", // "module",
                Name = "codeFragment",
            };
            using(var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(vb6Code)))
            using (var streamReader = new StreamReader(memoryStream))
            {
                ParseProcedures(streamReader);
                
                if(SourceModule.PropertyList.IsNotEmpty())
                    ConvertSource.MapClassProperties(SourceModule.PropertyList, TargetModule.PropertyList, TargetModule);

                if (SourceModule.ProcedureList.IsNotEmpty())
                    ConvertSource.Procedures(SourceModule, TargetModule.ProcedureList);
                
                var convertedCode = ConvertProcedures();
                if (convertedCode.IsEmpty())
                    convertedCode = GenerateProperties();

                var translatedCode = Massage.BlanketReplaceNow(convertedCode, SourceModule.Name);
                return translatedCode;
            }

            return "";
        }

        public string GenerateCode(string inputFilename, string outputPath = null)
        {
            var version = string.Empty;

            // try recognize source code type depend by file extension
            var extension = inputFilename.Substring(inputFilename.Length - 3, 3);
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
            using (var stream = new FileStream(inputFilename, FileMode.Open, FileAccess.Read))
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
                    return null;
                }

                SourceModule = new Module(FirstParent)
                {
                    Version = version ?? "1.0",
                    FileName = inputFilename
                };

                // now parse specifics of each type
                switch (extension.ToUpper())
                {
                    case "FRM":
                        SourceModule.Type = "form";
                        ParseForm(reader);
                        break;

                    case "BAS":
                        SourceModule.Type = "module";
                        ParseModule(reader);
                        break;

                    case "CLS":
                        SourceModule.Type = "class";
                        ParseClass(reader);
                        break;
                }

                // parse remain - variables, functions, procedures
                ParseProcedures(reader);

                stream.Flush();
                stream.Close();
                reader.Close();
            }

            // generate output file
            var code = GetModuleCode(outputPath);
            return code;
        }

        public void ConvertFormCode(StringBuilder result, string outputPath = null)
        {
            // list of controls
            foreach (var control in TargetModule.ControlList)
            {
                if (!control.Valid) result.Append("//");

                result.Append(Tools.Indent(2) + " public System.Windows.Forms." + control.Type +
                              " " + control.Name + ";\r\n");
            }

            result.Append(Indent2 + "/// <summary>\r\n");
            result.Append(Indent2 + "/// Required designer variable.\r\n");
            result.Append(Indent2 + "/// </summary>\r\n");
            result.Append(Indent2 +
                          "public System.ComponentModel.Container components = null;\r\n");
            result.AppendLine();
            result.Append(Indent2 + "public " + SourceModule.Name + "()\r\n");
            result.Append(Indent2 + "{\r\n");
            result.Append(Indent3 + "// Required for Windows Form Designer support\r\n");
            result.Append(Indent3 + "InitializeComponent();\r\n");
            result.AppendLine();
            result.Append(Indent3 +
                          "// TODO: Add any constructor code after InitializeComponent call\r\n");
            result.Append(Indent2 + "}\r\n");

            result.Append(Indent2 + "/// <summary>\r\n");
            result.Append(Indent2 + "/// Clean up any resources being used.\r\n");
            result.Append(Indent2 + "/// </summary>\r\n");
            result.Append(Indent2 + "protected override void Dispose( bool disposing )\r\n");
            result.Append(Indent2 + "{\r\n");
            result.Append(Indent3 + "if( disposing )\r\n");
            result.Append(Indent3 + "{\r\n");
            result.Append(Indent3 + "  if (components != null)\r\n");
            result.Append(Indent3 + "  {\r\n");
            result.Append(Indent3 + "    components.Dispose();\r\n");
            result.Append(Indent3 + "  }\r\n");
            result.Append(Indent3 + "}\r\n");
            result.Append(Indent3 + "base.Dispose( disposing );\r\n");
            result.Append(Indent2 + "}\r\n");

            result.Append(Indent2 + "#region Windows Form Designer generated code\r\n");
            result.Append(Indent2 + "/// <summary>\r\n");
            result.Append(Indent2 + "/// Required method for Designer support - do not modify\r\n");
            result.Append(Indent2 + "/// the contents of this method with the code editor.\r\n");
            result.Append(Indent2 + "/// </summary>\r\n");
            result.Append(Indent2 + "public void InitializeComponent()\r\n");
            result.Append(Indent2 + "{\r\n");

            // if form contain images
            if (TargetModule.ImagesUsed)
                // System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
                result.Append(Indent3 + "System.Resources.ResourceManager resources = " +
                              "new System.Resources.ResourceManager(typeof(" + TargetModule.Name +
                              "));\r\n");

            foreach (var control in TargetModule.ControlList)
            {
                if (!control.Valid) result.Append("//");

                result.Append(Indent3 + "this." + control.Name
                              + " = new System.Windows.Forms." + control.Type
                              + "();\r\n");
            }

            // SuspendLayout part
            result.Append(Indent3 + "this.SuspendLayout();\r\n");
            // this.Frame1.ResumeLayout(false);
            // resume layout for each container
            foreach (var control in TargetModule.ControlList)
                // check if control is container
                // !! for menu controls
                if (control.Container && control.Type != "MenuItem" && control.Type != "MainMenu")
                {
                    if (!control.Valid) result.Append("//");

                    result.Append(Indent3 + "this." + control.Name + ".SuspendLayout();\r\n");
                }

            // each controls and his property
            foreach (var control in TargetModule.ControlList)
            {
                result.Append(Indent3 + "//\r\n");
                result.Append(Indent3 + "// " + control.Name + "\r\n");
                result.Append(Indent3 + "//\r\n");

                // unsupported control
                if (!control.Valid) result.Append("/*");

                // ImageList, Timer, Menu has't name property
                if (control.Type != "ImageList" && control.Type != "Timer"
                                                && control.Type != "MenuItem" &&
                                                control.Type != "MainMenu")
                    // control name
                    result.Append(Indent3 + "this." + control.Name + ".Name = "
                                  + (char) 34 + control.Name + (char) 34 + ";\r\n");

                // write properties
                foreach (var property in control.Children.Cast<ControlProperty>())
                    GetPropertyRow(result, control.Type, control.Name, property, outputPath);

                // if control is container for other controls
                var temp = string.Empty;
                foreach (var oControl1 in TargetModule.ControlList)
                    // all controls ownered by current control
                    if (oControl1.Owner == control.Name && !oControl1.InvisibleAtRuntime)
                        temp += Indent3 + Indent3 + "this." + oControl1.Name + ",\r\n";

                if (temp != string.Empty)
                {
                    // exception for menu controls
                    if (control.Type == "MainMenu" || control.Type == "MenuItem")
                        // this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                        result.Append(Indent3 + "this." + control.Name
                                      + ".MenuItems.AddRange(new System.Windows.Forms.MenuItem[]\r\n");
                    else
                        // this. + control.Name + .Controls.AddRange(new System.Windows.Forms.Control[]
                        result.Append(Indent3 + "this." + control.Name
                                      + ".Controls.AddRange(new System.Windows.Forms.Control[]\r\n");

                    result.Append(Indent3 + "{\r\n");
                    result.Append(temp);
                    // remove last comma, keep CRLF
                    result.Remove(result.Length - 3, 1);
                    // close addrange part
                    result.Append(Indent3 + "});\r\n");
                }

                // unsupported control
                if (!control.Valid) result.Append("*/");
            }

            result.Append(Indent3 + "//\r\n");
            result.Append(Indent3 + "// " + SourceModule.Name + "\r\n");
            result.Append(Indent3 + "//\r\n");
            result.Append(Indent3 +
                          "this.Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
            result.Append(Indent3 + "{\r\n");

            // add control range to form
            foreach (var control in TargetModule.ControlList)
            {
                if (!control.Valid) result.Append("//");

                // all controls ownered by main form
                if (control.Owner == SourceModule.Name && !control.InvisibleAtRuntime)
                    result.Append(Indent3 + "      this." + control.Name + ",\r\n");
            }

            // remove last comma, keep CRLF
            result.Remove(result.Length - 3, 1);
            // close addrange part
            result.Append(Indent3 + "});\r\n");

            // form name
            result.Append(Indent3 + "this.Name = " + (char) 34 + TargetModule.Name + (char) 34 +
                          ";\r\n");
            // exception for menu
            // this.Menu = this.mainMenu1;
            if (TargetModule.MenuUsed)
                foreach (var control in TargetModule.ControlList)
                    if (control.Type == "MainMenu")
                        result.Append(Indent3 + "      this.Menu = " + control.Name + ";\r\n");

            // form properties
            foreach (var property in TargetModule.FormPropertyList)
            {
                if (!property.Valid) result.Append("//");

                GetPropertyRow(result, TargetModule.Type, "", property, outputPath);
            }

            // resume layout for each container
            foreach (var control in TargetModule.ControlList)
                // check if control is container
                if (control.Container && control.Type != "MenuItem" && control.Type != "MainMenu")
                {
                    if (!control.Valid) result.Append("//");

                    result.Append(Indent3 + "this." + control.Name + ".ResumeLayout(false);\r\n");
                }

            // form
            result.Append(Indent3 + "this.ResumeLayout(false);\r\n");

            result.Append(Indent2 + "}\r\n");
            result.Append(Indent2 + "#endregion\r\n");
        }

        public string ConvertProcedures()
        {
            var result = new StringBuilder();
            foreach (var procedure in TargetModule.ProcedureList)
                result.AppendLine(procedure.GenerateCode());

            return result.ToString();
        }

        public string GenerateProperties()
        {
            if (TargetModule.PropertyList.Count <= 0)
                return string.Empty;

            var result = new StringBuilder();

            result.AppendLine();
            foreach (var property in TargetModule.PropertyList)
            {
                property.TargetModule = TargetModule;
                property.ResetIndent(TargetModule.Indent + 2);

                result.AppendLine(
                    // firstIndent +
                    Massage.AllLinesNow(property.Final, SourceModule)
                );
            }

            var code = result.ToString();
            return code;
        }

        public string ConvertVariablesInCode()
        {
            var code = TargetModule.VariableList.GenerateCode(TargetModule.Indent + 2);
            return code;
        }

        public string GetModuleCode(string outputPath = null)
        {
            var result = new StringBuilder();

            // convert source to target
            if (TargetModule == null)
                TargetModule = new Module(FirstParent);

            ConvertSource.Module(SourceModule, TargetModule);

            // ********************************************************
            // common class
            // ********************************************************
            result.Append("using System;\r\n");
            result.Append("using System.Collections;\r\n");
            result.Append("using System.Collections.Generic;\r\n");

            // ********************************************************
            // only form class
            // ********************************************************
            if (TargetModule.Type == "form")
            {
                result.Append("using System.Drawing;\r\n");
                result.Append("using System.ComponentModel;\r\n");
                result.Append("using System.Windows.Forms;\r\n");
            }

            result.AppendLine();
            result.AppendLine($"namespace {ProjectNamespace}");
            // start namespace region
            result.Append("{\r\n");
            var firstIndent = Tools.Indent(1);

            if (SourceModule.Comment.IsEmpty())
                SourceModule.Comment = SourceModule.PropertyList?.FirstOrDefault()?.Comment
                                       ?? SourceModule.ProcedureList?.FirstOrDefault()?.Comment;

            if (SourceModule.Comment.IsNotEmpty())
            {
                result.Append($"{firstIndent}/// <summary>\r\n");
                result.Append($"{firstIndent}///   {SourceModule.Comment}.\r\n");
                result.Append($"{firstIndent}/// </summary>\r\n");
            }

            switch (TargetModule.Type)
            {
                case "form":
                    result.AppendLine(
                        $"{firstIndent}public class {SourceModule.Name} : System.Windows.Forms.Form");
                    break;

                case "module":
                    result.AppendLine($"{firstIndent}class {SourceModule.Name}");
                    // all procedures must be static
                    break;

                case "class":
                    //if (SourceModule.Name.ToUpper().StartsWith("I"))
                    //    TargetModule.Type = "interface";

                    result.AppendLine(
                        $"{firstIndent}public {TargetModule.Type} {SourceModule.Name}");
                    break;
            }

            // start class region
            result.AppendLine($"{firstIndent}{{");

            // ********************************************************
            // only form class
            // ********************************************************

            if (TargetModule.Type == "form")
                ConvertFormCode(result, outputPath);

            // ********************************************************
            // enums
            // ********************************************************

            if (TargetModule.EnumList.Count > 0)
                result.AppendLine(ConvertEnumCode(3));

            // ********************************************************
            //  variables for al module types
            // ********************************************************

            if (TargetModule.VariableList.Count > 0)
                result.AppendLine(ConvertVariablesInCode()); // 3));

            // ********************************************************
            // properties has only forms and classes
            // ********************************************************

            if (TargetModule.Type == "form" || TargetModule.Type == "class")
                result.AppendLine(GenerateProperties()); // 2));

            // ********************************************************
            // procedures
            // ********************************************************

            if (TargetModule.ProcedureList.Count > 0)
                result.Append(ConvertProcedures()); // 2));

            // end class
            result.AppendLine($"{firstIndent}}}");

            // end namespace
            result.AppendLine("    }");

            var code = Massage.BlanketReplaceNow(result.ToString(), SourceModule.Name);
            return code;
        }

        public void GetPropertyRow(StringBuilder result, string type, string name,
            ControlProperty controlProperty,
            string outputPath = null)
        {
            // exception for images
            if (controlProperty.Name == "Icon"
                || controlProperty.Name == "Image"
                || controlProperty.Name == "BackgroundImage")
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

                if (!WriteImage(SourceModule, resourceName, controlProperty.Value, outputPath))
                    return;

                switch (controlProperty.Name)
                {
                    case "BackgroundImage":
                        result.Append(Indent3 + "this."
                                              + controlProperty.Name +
                                              " = ((System.Drawing.Bitmap)(resources.GetObject("
                                              + (char) 34 + "$this.BackgroundImage" + (char) 34 +
                                              ")));\r\n");
                        break;

                    case "Icon":
                        result.Append(Indent3 + "this."
                                              + controlProperty.Name +
                                              " = ((System.Drawing.Icon)(resources.GetObject("
                                              + (char) 34 + "$this.Icon" + (char) 34 + ")));\r\n");
                        break;

                    case "Image":
                        result.Append(Indent3 + "this." + name + "."
                                      + controlProperty.Name +
                                      " = ((System.Drawing.Bitmap)(resources.GetObject("
                                      + (char) 34 + name + ".Image" + (char) 34 + ")));\r\n");
                        break;
                }
            }
            else
            {
                // unsupported property
                if (!controlProperty.Valid) result.Append("//");

                if (type == "form")
                    // form properties
                    result.Append(Indent3 + "this."
                                          + controlProperty.Name + " = " + controlProperty.Value +
                                          ";\r\n");
                else
                    // control properties
                    result.Append(Indent3 + "this." + name + "."
                                  + controlProperty.Name + " = " + controlProperty.Value + ";\r\n");
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
                            SourceModule.Name =
                                GetWord(line, ref position).Replace("\"", string.Empty);
                            break;

                        case "VB_Exposed":
                            return true;

                        case "VB_Description":
                            position++;
                            tempString = GetWord(line, ref position);
                            position++;
                            SourceModule.Comment =
                                GetWord(line, ref position).Replace("\"", string.Empty);
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
            string sType;
            var iPosition = 0;
            var iLevel = 0;
            Control control = null;
            ControlProperty oNestedProperty = null;
            var bNestedProperty = false;

            // parse only visual part of form
            while (!bFinish)
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
                                SourceModule.ControlList.Add(control);
                            }

                            // save name of previous control as owner for current and next controls
                            OwnerStock.Add(sOwner);
                        }

                        bEnd = false;

                        switch (sType)
                        {
                            case "Form": // VERSION 2.00 - VB3
                            case "VB.Form":
                            case "VB.MDIForm":
                                SourceModule.Name = sName;
                                // first owner
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                break;

                            default:
                                if (TargetModule == null)
                                    TargetModule = new Module(FirstParent);

                                // new control
                                control = new Control(TargetModule)
                                {
                                    Name = sName,
                                    Type = sType
                                };
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                // set current container name
                                control.Owner = OwnerStock[OwnerStock.Count - 1];
                                break;
                        }

                        break;

                    case "End":
                        // double end - we leaving some container
                        if (bEnd)
                        {
                            // remove last item from stock
                            OwnerStock.Remove(OwnerStock[OwnerStock.Count - 1]);
                        }
                        else
                        {
                            // level 1 is form and all higher levels are controls
                            if (iLevel > 1)
                                // add control to colection
                                SourceModule.ControlList.Add(control);
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

                        oNestedProperty = new ControlProperty(control, 1);
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
                            // add property to form
                            SourceModule.FormPropertyList.Add(oNestedProperty);
                        else
                            // to controls
                            control.Children.Add(oNestedProperty);

                        break;

                    default:
                        // parse property
                        var property = new ControlProperty(FirstParent, 2);

                        var iTemp = sLine.IndexOf("=");
                        if (iTemp > -1)
                        {
                            property.Name = sLine.Substring(0, iTemp - 1).Trim();
                            iComment = sLine.IndexOf("'", iTemp);
                            if (iComment > -1)
                            {
                                property.Value = sLine.Substring(iTemp + 1, iComment - iTemp - 1)
                                    .Trim();
                                property.Comment =
                                    sLine.Substring(iComment + 1, sLine.Length - iComment - 1)
                                        .Trim();
                            }
                            else
                            {
                                property.Value = sLine
                                    .Substring(iTemp + 1, sLine.Length - iTemp - 1).Trim();
                            }

                            if (bNestedProperty)
                            {
                                oNestedProperty.PropertyList.Add(property);
                            }
                            else
                            {
                                // depend by level insert property to form or control
                                if (iLevel > 1)
                                    // add property to control
                                    control.Children.Add(property);
                                else
                                    // add property to form
                                    SourceModule.FormPropertyList.Add(property);
                            }
                        }

                        break;
                }

                if (iLevel == 0 && bProcess)
                    // visual part of form is finish
                    bFinish = true;
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
            // parameters delimited by comma
            var parts = line
                .Split(',')
                .ToList()
                .Select(p => p.Trim())
                .ToList();

            foreach (var part in parts)
            {
                var parameter = new Parameter
                {
                    IsOptional = part.StartsWith("Optional")
                };

                var words = part.Replace("Optional", "").Trim();

                parameter.Pass = words.StartsWith("ByVal") ? "ByVal" : "ByRef";
                words = words.Replace("ByRef ", "");
                words = words.Replace("ByVal ", "").Trim();

                parameter.Name = words.TokenAt(1, " As ").Trim();

                var type = words.TokenAt(2, " As ").Trim().TokenAt(1).Trim();

                parameter.Type = type; //"/* unknown */";
                if (type.IsNotEmpty())
                {
                    var left = words
                        .TokenAt(2, " As ")
                        .Replace(type, "");
                    if (left.IsNotEmpty()) parameter.DefaultValue = left.Replace("= ", "");
                }

                if (parameter.Type == "String")
                    parameter.Type = "string";

                parameterList.Add(parameter);
            }
        }

        // ByVal lValue As Long, ByVal sValue As string
        public void ParseProcedureName(Procedure procedure, string line)
        {
            var word = string.Empty;
            var position = 0;
            var start = 0;

            word = GetWord(line, ref position);
            switch (word)
            {
                case "Private":
                    procedure.Scope = "private";
                    break;

                case "Public":
                    procedure.Scope = "public";
                    break;

                default:
                    procedure.Scope = "private";
                    break;
            }

            // property
            position++;
            word = GetWord(line, ref position);

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

                default:
                    throw new NotSupportedException();
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

            if (position - start > 0)
            {
                var parameters = line.Substring(start, position - start);
                ParseParameters(procedure.ParameterList, parameters);
            }

            // and return type of function
            if (procedure.Type == ProcedureType.ProcedureFunction)
            {
                // as
                position++;
                word = GetWord(line, ref position);
                // function return type
                position++;
                var @as = GetWord(line, ref position);
                procedure.ReturnType = @as == "As" 
                    ? line.LastToken("As").Trim() 
                    : GetWord(line, ref position);
            }
        }

        public bool ParseProcedures(StreamReader reader)
        {
            string sComments = null;
            var insideAnEnum = false;
            var insideAVariable = false;
            var insideAProperty = false;
            var insideAProcedure = false;
            var atEndOfModule = false;

            Variable variable = null;

            Property vb6Property = null;
            Procedure vb6Procedure = null;
            Enum vb6EnumItems = null;

            if (TargetModule == null)
                TargetModule = new Module(FirstParent);

            while (reader.Peek() > -1)
            {
                var line = reader.ReadLine();
                var iPosition = 0;
                if (line.IsEmpty())
                    continue;

                // check if next line is same command, join it together ?
                //while (line.Substring(line.Length - 1, 1) == "_") line += reader.ReadLine();
                while (line.EndsWith("_"))
                    line += reader.ReadLine();

                // get first word in line
                var word = GetWord(line, ref iPosition);
                if (int.TryParse(word, out var tempResult))
                {
                    line = line.TokensAfter().Trim();
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
                        foreach (var commentLine in line.Lines(
                            StringSplitOptions.RemoveEmptyEntries))
                            sComments += "// "
                                         + (commentLine.Left(1) == "'"
                                             ? commentLine.Substring(1)
                                             : commentLine)
                                         + Environment.NewLine;
                        break;

                    case "Public":
                    case "Private":
                        // save it for later use
                        var sScope = word.ToLower();
                        // read next word
                        // next word - control type
                        iPosition++;
                        word = GetWord(line, ref iPosition);

                        switch (word)
                        {
                            // functions or procedures
                            case "Sub":
                            case "Function":
                                vb6Procedure = new Procedure();
                                vb6Procedure.Comment = sComments;
                                sComments = string.Empty;
                                ParseProcedureName(vb6Procedure, line);
                                insideAProcedure = true;
                                break;

                            case "Enum":
                                vb6EnumItems = new Enum();
                                vb6EnumItems.Scope = sScope;
                                // next word is enum name
                                iPosition++;
                                vb6EnumItems.Name = GetWord(line, ref iPosition);
                                insideAnEnum = true;
                                break;

                            case "Property":
                                vb6Property = new Property(TargetModule, sComments);
                                sComments = string.Empty;
                                ParsePropertyName(vb6Property, line);
                                insideAProperty = true;
                                break;

                            default:
                                // variable declaration
                                variable = new Variable();
                                ParseVariableDeclaration(variable, line);
                                insideAVariable = true;
                                break;
                        }

                        break;

                    case "Dim":
                        // variable declaration
                        variable = new Variable();
                        variable.Comment = sComments;
                        sComments = string.Empty;
                        ParseVariableDeclaration(variable, line);
                        insideAVariable = true;
                        break;

                    // functions or procedures
                    case "Sub":
                    case "Function":
                        break;

                    case "End":
                        atEndOfModule = true;
                        break;

                    default:
                        if (insideAnEnum)
                        {
                            // first word is name, second =, third value if is preset
                            var vb6EnumItem = new EnumItem {Comment = sComments};
                            sComments = string.Empty;
                            ParseEnumItem(vb6EnumItem, line);
                            // add item
                            vb6EnumItems.ItemList.Add(vb6EnumItem);
                        }

                        if (insideAProperty)
                        {
                            line = line.Trim();
                            if (vb6Property.Direction.ToLower() == "get"
                                && line.FirstToken(" = ").ToLower() == vb6Property.Name.ToLower()
                            )
                                vb6Property.Block.Children.Add(Quick.Line(vb6Property,
                                    "return " + line.TokensAfterFirst(" = ")));
                            else if (vb6Property.Direction.ToLower() == "set" ||
                                     vb6Property.Direction.ToLower() == "let"
                                     && line.FirstToken(" = ").ToLower() ==
                                     vb6Property.Name.ToLower()
                            )
                                vb6Property.Line = "return " + line.TokensAfterFirst(" = ");
                            else
                                // add line of property
                                vb6Property.Block.Children.Add(Quick.Line(vb6Property, line));
                        }

                        if (insideAProcedure) 
                            vb6Procedure.LineList.Add(line);

                        break;
                }

                // if something end
                if (atEndOfModule)
                {
                    if (insideAnEnum)
                    {
                        SourceModule.EnumList.Add(vb6EnumItems);
                        insideAnEnum = false;
                    }

                    if (insideAProperty)
                    {
                        SourceModule.PropertyList.Add(vb6Property);
                        insideAProperty = false;
                    }

                    if (insideAProcedure)
                    {
                        SourceModule.ProcedureList.Add(vb6Procedure);
                        insideAProcedure = false;
                    }

                    atEndOfModule = false;
                }
                else
                {
                    if (insideAVariable) 
                        SourceModule.VariableList.Add(variable);
                }
                insideAVariable = false;
            }
            return true;
        }

        public void ParsePropertyName(IAmAProperty property, string line)
        {
            var localProperty = (Property) property;

            var iPosition = 0;
            var start = 0;
            var status = false;

            // next word - control type
            var parameters = GetWord(line, ref iPosition);
            switch (parameters)
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
                parameters = GetWord(line, ref iPosition);
            }

            // direction Let, Get, Set
            iPosition++;
            parameters = GetWord(line, ref iPosition);
            localProperty.Direction = parameters;

            //Public Property Let ParentID(ByVal lValue As Long)

            // name
            start = iPosition;
            iPosition = line.IndexOf("(", start + 1);
            property.Name = line.Substring(start, iPosition - start).Trim();

            // + possible parameters
            iPosition++;
            start = iPosition;
            iPosition = line.IndexOf(")", start);

            if (iPosition - start > 0)
            {
                parameters = line.Substring(start, iPosition - start);
                // process parameters
                ParseParameters(localProperty.Parameters, parameters);
            }

            // As
            iPosition++;
            iPosition++;
            parameters = GetWord(line, ref iPosition);

            // type
            iPosition++;
            parameters = GetWord(line, ref iPosition);
            if (parameters.IsNotEmpty())
                property.Type = parameters;
        }

        public void ParseVariableDeclaration(Variable variable, string line)
        {
            var word = string.Empty;
            var iPosition = 0;
            var status = false;
            line = line.Deflate();

            // next word - control type
            word = GetWord(line, ref iPosition);
            switch (word)
            {
                case "Dim":
                case "Private":
                    variable.Scope = "public"; // "private";
                    break;

                case "Public":
                    variable.Scope = "public";
                    break;

                default:
                    variable.Scope = "public"; // "private";
                    variable.Name = variable.Scope;
                    status = true;
                    break;
            }

            // variable name
            if (!status)
            {
                iPosition++;
                word = GetWord(line, ref iPosition);
                variable.Name = word;
            }

            // As
            iPosition++;
            word = GetWord(line, ref iPosition);
            // variable type
            iPosition++;
            word = GetWord(line, ref iPosition);
            
            // variable.Type = word.CSharpTypeEquivalent();
            variable.Type = word == "String" ? "string" : word;
        }

        public bool WriteImage(Module sourceModule, string resourceName, string value,
            string outputPath = null)
        {
            var temp = string.Empty;
            var offset = 0;
            var frxFile = string.Empty;
            var sResxName = string.Empty;
            var position = 0;

            position = value.IndexOf(":", 0, StringComparison.Ordinal);

            // "Form1.frx":0000;
            // old vb3 code has name without ""
            // CONTROL.FRX:0000

            if (sourceModule.Version == "5.00")
                frxFile = value.Substring(1, position - 2);
            else
                frxFile = value.Substring(0, position);

            temp = value.Substring(position + 1, value.Length - position - 1);
            offset = Convert.ToInt32("0x" + temp, 16);
            // exist file ?

            // get image
            Tools.GetFrxImage(Path.GetDirectoryName(sourceModule.FileName) + @"\" + frxFile, offset,
                out var imageString);

            if (!string.IsNullOrEmpty(outputPath)
                && imageString.GetLength(0) - 8 > 0)
            {
                var resourcePath = Path.Combine(outputPath, resourceName);
                if (File.Exists(resourcePath))
                {
                    File.SetAttributes(resourcePath, FileAttributes.Normal);
                    File.Delete(resourcePath);
                }

                var stream = new FileStream(resourcePath, FileMode.CreateNew, FileAccess.Write);
                using (var writer = new BinaryWriter(stream))
                {
                    writer.Write(imageString, 8, imageString.GetLength(0) - 8);
                    stream.Close();
                    writer.Close();
                }

                TargetModule.ImageList.Add(resourceName);
                return true;
            }

            return false;
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
                    try
                    {
                        var img = Image.FromFile(outPath + resourceName);
                        rsxw.AddResource(resourceName, img);
                        img.Dispose();
                    }
                    catch
                    {
                    }

                // rsxw.Generate();
                rsxw.Close();

                foreach (var resourceName in mImageList) File.Delete(outPath + resourceName);
            }
        }
    }
}