using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Xml;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    // parse VB6 properties and values to C#
    internal sealed class Tools
    {
        private static Hashtable _mControlList;

        public static void GetFrxImage(string imageFile, int imageOffset, out byte[] imageString)
        {
            byte[] header;
            var bytesToRead = 0;

            // open file
            var stream = new FileStream(imageFile, FileMode.Open, FileAccess.Read);
            var reader = new BinaryReader(stream);
            // Start from offset
            reader.BaseStream.Seek(imageOffset, SeekOrigin.Begin);
            // Get the four byte header
            header = new byte[4];
            header = reader.ReadBytes(4);
            // Convert This Header Into The Number Of Bytes
            // To Read For This Image
            bytesToRead = header[0];
            bytesToRead = bytesToRead + (header[1] * 0x100);
            bytesToRead = bytesToRead + (header[2] * 0x10000);
            bytesToRead = bytesToRead + (header[3] * 0x1000000);
            // Get image information
            imageString = new byte[bytesToRead];
            imageString = reader.ReadBytes(bytesToRead);

            //      Stream = new FileStream( @"C:\temp\test\Ba.bmp", FileMode.CreateNew, FileAccess.Write );
            //      BinaryWriter Writer = new BinaryWriter( Stream );
            //      Writer.Write(ImageString, 8, ImageString.GetLength(0) - 8);
            //      Stream.Close();
            //      Writer.Close();

            //  FileStream inFile = new FileStream(@"C:\WINdows\Blue Lace 16.bmp", FileMode.Open, FileAccess.Read);
            //  ReturnImage = Image.FromStream(inFile, false);

            stream.Close();
            reader.Close();
        }

        public static bool ParseClassProperties(Module sourceModule, Module targetModule)
        {
            foreach (Property sourceProperty in sourceModule.PropertyList)
            {
                var targetProperty = new Property
                {
                    Name = sourceProperty.Name, 
                    Comment = sourceProperty.Comment,
                    Scope = sourceProperty.Scope,
                    Type = VariableTypeConvert(sourceProperty.Type),
                };

                switch (sourceProperty.Direction)
                {
                    case "Get":
                        targetProperty.Direction = "get";
                        break;

                    case "Set":
                    case "Let":
                        targetProperty.Direction = "set";
                        break;
                }
                // lines
                foreach (string line in sourceProperty.LineList)
                    if (line.Trim() != string.Empty)
                        targetProperty.LineList.Add(line);

                targetModule.PropertyList.Add(targetProperty);
            }
            return true;
        }

        public static bool ParseControlProperties(Module oModule, Control oControl,
                                                ArrayList sourcePropertyList,
                                                ArrayList targetPropertyList)
        {
            ControlProperty targetProperty = null;

            // each property
            foreach (ControlProperty sourceProperty in sourcePropertyList)
            {
                if (sourceProperty.Name == "Index")
                {
                    // Index           =   3
                    oControl.Name = oControl.Name + sourceProperty.Value;
                }
                else
                {
                    targetProperty = new ControlProperty();
                    if (ParseProperties(oControl.Type, sourceProperty, targetProperty, sourcePropertyList))
                    {
                        if (targetProperty.Name == "Image")
                        {
                            oModule.ImagesUsed = true;
                        }
                        targetPropertyList.Add(targetProperty);
                    }
                }
            }
            return true;
        }

        public static bool ParseControls(Module oModule, ArrayList sourceControlList, ArrayList targetControlList)
        {
            var type = string.Empty;

            foreach (Control oSourceControl in sourceControlList)
            {
                var oTargetControl = new Control();

                oTargetControl.Name = oSourceControl.Name;
                oTargetControl.Owner = oSourceControl.Owner;
                oTargetControl.Container = oSourceControl.Container;
                oTargetControl.Valid = true;

                // compare upper case type
                if (_mControlList.ContainsKey(oSourceControl.Type.ToUpper()))
                {
                    var oItem = (ControlListItem)_mControlList[oSourceControl.Type.ToUpper()];

                    if (oItem.Unsupported)
                    {
                        type = "Unsuported";
                        oTargetControl.Valid = false;
                    }
                    else
                    {
                        type = oItem.CsharpName;
                        if (type == "MenuItem")
                        {
                            oModule.MenuUsed = true;
                        }
                    }
                    oTargetControl.InvisibleAtRuntime = oItem.InvisibleAtRuntime;
                }
                else
                {
                    type = oSourceControl.Type;
                }

                oTargetControl.Type = type;
                ParseControlProperties(oModule, oTargetControl, oSourceControl.PropertyList, oTargetControl.PropertyList);

                targetControlList.Add(oTargetControl);
            }
            return true;
        }

        public static bool ParseEnums(Module sourceModule, Module targetModule)
        {
            foreach (Enum sourceEnum in sourceModule.EnumList)
            {
                targetModule.EnumList.Add(sourceEnum);
            }
            return true;
        }

        public static bool ParseModule(Module sourceModule, Module targetModule)
        {
            ControlListLoad();

            // module name
            targetModule.Name = sourceModule.Name;
            // file name
            targetModule.FileName = Path.GetFileNameWithoutExtension(sourceModule.FileName) + ".cs";
            // type
            targetModule.Type = sourceModule.Type;
            // version
            targetModule.Version = sourceModule.Version;
            // process own properties - forms
            ParseModuleProperties(targetModule, sourceModule.FormPropertyList, targetModule.FormPropertyList);
            // process controls - form
            ParseControls(targetModule, sourceModule.ControlList, targetModule.ControlList);

            // special exception for menu
            if (targetModule.MenuUsed)
            {
                // add main menu control
                var oControl = new Control
                {
                    Name = "MainMenu",
                    Owner = targetModule.Name,
                    Type = "MainMenu",
                    Valid = true,
                    InvisibleAtRuntime = true
                };
                targetModule.ControlList.Insert(0, oControl);
                foreach (Control oMenuControl in targetModule.ControlList)
                    if ((oMenuControl.Type == "MenuItem") && (oMenuControl.Owner == targetModule.Name))
                        // rewrite previous owner
                        oMenuControl.Owner = oControl.Name;
            }

            var tempControlList = new ArrayList();
            var tabControlIndex = 0;

            // check for TabDlg.SSTab
            foreach (Control oTargetControl in targetModule.ControlList)
            {
                if ((oTargetControl.Type == "TabControl") && (oTargetControl.Valid))
                {
                    // for each source table is necessary
                    //          this.tabControl1 = new System.Windows.Forms.TabControl();
                    //          this.tabPage1 = new System.Windows.Forms.TabPage();

                    var index = 0;
                    Control oTabPage = null;
                    // each property
                    foreach (ControlProperty oTargetProperty in oTargetControl.PropertyList)
                    {
                        // TabCaption = create new tab
                        //      this.SSTab1.(TabCaption(0)) = "Tab 0";

                        Console.WriteLine(oTargetProperty.Name);

                        if (oTargetProperty.Name.IndexOf("TabCaption(" + index.ToString() + ")", 0) > -1)
                        {
                            // new tab
                            oTabPage = new Control();
                            oTabPage.Type = "TabPage";
                            oTabPage.Name = "tabPage" + index.ToString();
                            oTabPage.Owner = oTargetControl.Name;
                            oTabPage.Container = true;
                            oTabPage.Valid = true;
                            oTabPage.InvisibleAtRuntime = false;

                            // add some necessary properties
                            var targetProperty = new ControlProperty();
                            targetProperty.Name = "Location";
                            targetProperty.Value = "new System.Drawing.Point(4, 22)";
                            targetProperty.Valid = true;
                            oTabPage.PropertyList.Add(targetProperty);

                            targetProperty = new ControlProperty();
                            targetProperty.Name = "Size";
                            targetProperty.Value = "new System.Drawing.Size(477, 374)";
                            targetProperty.Valid = true;
                            oTabPage.PropertyList.Add(targetProperty);

                            targetProperty = new ControlProperty();
                            targetProperty.Name = "Text";
                            targetProperty.Value = oTargetProperty.Value;
                            targetProperty.Valid = true;
                            oTabPage.PropertyList.Add(targetProperty);

                            targetProperty = new ControlProperty();
                            targetProperty.Name = "TabIndex";
                            targetProperty.Value = index.ToString();
                            targetProperty.Valid = true;
                            oTabPage.PropertyList.Add(targetProperty);

                            tempControlList.Add(oTabPage);
                            index++;
                        }

                        // Control = change owner of control to current tab
                        //      this.SSTab1.(Tab(0).Control(0) = "ImageControl";
                        if (oTargetProperty.Name.IndexOf(".Control(", 0) > -1)
                        {
                            if (oTargetProperty.Name.IndexOf("Enable", 0) == -1)
                            {
                                var tabName = oTargetProperty.Value.Substring(1, oTargetProperty.Value.Length - 2);
                                tabName = GetControlIndexName(tabName);
                                // search for "oTargetProperty.Value" control
                                // and replace owner of this control to current tab
                                foreach (Control oNewOwner in targetModule.ControlList)
                                {
                                    if ((oNewOwner.Name == tabName) && (!oNewOwner.InvisibleAtRuntime))
                                    {
                                        oNewOwner.Owner = oTabPage.Name;
                                    }
                                }
                            }
                        }
                    }
                }
                tabControlIndex++;
            }

            if (tempControlList.Count > 0)
            {
                // right order of tabs
                var position = 0;
                foreach (Control oControl in tempControlList)
                {
                    targetModule.ControlList.Insert(tabControlIndex + position, oControl);
                    position++;
                }
            }

            // process enums
            ParseEnums(sourceModule, targetModule);
            // process variables
            ParseVariables(sourceModule.VariableList, targetModule.VariableList);
            // process properties
            ParseClassProperties(sourceModule, targetModule);
            // process procedures
            ParseProcedures(sourceModule, targetModule);

            return true;
        }

        public static bool ParseModuleProperties(Module oModule,
                                              ArrayList sourcePropertyList,
                                              ArrayList targetPropertyList)
        {
            // each property
            foreach (ControlProperty sourceProperty in sourcePropertyList)
            {
                var targetProperty = new ControlProperty();
                if (ParseProperties(oModule.Type, sourceProperty, targetProperty, sourcePropertyList))
                {
                    if (targetProperty.Name == "BackgroundImage" || targetProperty.Name == "Icon")
                    {
                        oModule.ImagesUsed = true;
                    }
                    targetPropertyList.Add(targetProperty);
                }
            }
            return true;
        }

        public static bool ParseProcedures(Module sourceModule, Module targetModule)
        {
            const string indent6 = "      ";

            foreach (Procedure sourceProcedure in sourceModule.ProcedureList)
            {
                var targetProcedure = new Procedure
                {
                    Name = sourceProcedure.Name,
                    Scope = sourceProcedure.Scope,
                    Comment = sourceProcedure.Comment,
                    Type = sourceProcedure.Type,
                    ReturnType = VariableTypeConvert(sourceProcedure.ReturnType),
                    ParameterList = sourceProcedure.ParameterList
                };
                
                // lines
                foreach (string line in sourceProcedure.LineList)
                {
                    var tempSource = line.Trim();
                    if (tempSource.Length > 0)
                    {
                        var translatedLine = string.Empty;
                        if (tempSource.IndexOf("vbNullString", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("vbNullString", "String.Empty");
                            tempSource = translatedLine;
                        }
                        // Nothing = null
                        if (tempSource.IndexOf("Nothing", 0) > -1)
                        {
                            translatedLine = tempSource
                                .Replace("Set ", "")
                                .Replace("Nothing", "null");
                            translatedLine += ";";
                            tempSource = translatedLine;
                        }
                        // Set
                        if (tempSource.IndexOf("Set ", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("Set ", " ");
                            tempSource = translatedLine;
                        }
                        // remark
                        if (tempSource[0] == '\'') // '
                        {
                            translatedLine = tempSource.Replace("'", "//");
                            tempSource = translatedLine;
                        }
                        // & to +
                        if (tempSource.IndexOf("&", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("&", "+");
                            tempSource = translatedLine;
                        }
                        // Select Case
                        if (tempSource.IndexOf("Select Case", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("Select Case", "switch");
                            tempSource = translatedLine;
                        }
                        // End Select
                        if (tempSource.IndexOf("End Select", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("End Select", "}");
                            tempSource = translatedLine;
                        }
                        // _
                        if (tempSource.IndexOf(" _", 0) > -1)
                        {
                            translatedLine = tempSource.Replace(" _", "\r\n");
                            tempSource = translatedLine;
                        }
                        // If
                        if (tempSource.IndexOf("If ", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("If ", "if ( ");
                            tempSource = translatedLine;
                        }
                        // Not
                        if (tempSource.IndexOf("Not ", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("Not ", "! ");
                            tempSource = translatedLine;
                        }
                        // then
                        if (tempSource.IndexOf(" Then", 0) > -1)
                        {
                            translatedLine = tempSource.Replace(" Then", " )\r\n" + indent6 + "{\r\n");
                            tempSource = translatedLine;
                        }
                        // else
                        if (tempSource.IndexOf("Else", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("Else", "}\r\n" + indent6 + "else\r\n" + indent6 + "{");
                            tempSource = translatedLine;
                        }
                        // End if
                        if (tempSource.IndexOf("End If", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("End If", "}");
                            tempSource = translatedLine;
                        }
                        // Unload Me
                        if (tempSource.IndexOf("Unload Me", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("Unload Me", "Close()");
                            tempSource = translatedLine;
                        }
                        // .Caption
                        if (tempSource.IndexOf(".Caption", 0) > -1)
                        {
                            translatedLine = tempSource.Replace(".Caption", ".Text");
                            tempSource = translatedLine;
                        }
                        // True
                        if (tempSource.IndexOf("True", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("True", "true");
                            tempSource = translatedLine;
                        }
                        // False
                        if (tempSource.IndexOf("False", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("False", "false");
                            tempSource = translatedLine;
                        }

                        // New
                        if (line.Contains("If ") 
                            && line.Contains("Then") 
                            && line.TokensAfter(1, "Then").Trim().Length > 0)
                        {
                            translatedLine = tempSource.Replace("New", "new");
                            tempSource = translatedLine;
                        }

                        // New
                        if (tempSource.IndexOf("New", 0) > -1)
                        {
                            translatedLine = tempSource.Replace("New", "new");
                            tempSource = translatedLine;
                        }

                        if (tempSource.IndexOf("On Error Resume Next", 0) > -1)
                        {
                            targetProcedure.BottomLineList.Add(@"
        catch(Exception e) 
        { 
            /* ON ERROR RESUME NEXT (ish) */ 
        }
");
                            tempSource = translatedLine = "try\r\n{\r\n";
                        }
                        else
                        {
                            targetProcedure.LineList.Add(translatedLine == string.Empty ? tempSource : translatedLine);
                        }
                    }
                    else
                    {
                        targetProcedure.LineList.Add(string.Empty);
                    }
                }

                targetModule.ProcedureList.Add(targetProcedure);
            }
            return true;
        }

        public static bool ParseProperties(string type,
                                        ControlProperty sourceProperty,
                                        ControlProperty targetProperty,
                                        ArrayList sourcePropertyList)
        {
            var validProperty = true;
            targetProperty.Valid = true;

            switch (sourceProperty.Name)
            {
                // not used
                case "Appearance":
                case "ScaleHeight":
                case "ScaleWidth":
                case "Style":             // button
                case "BackStyle":         //label
                case "IMEMode":
                case "WhatsThisHelpID":
                case "Mask":              // maskedit
                case "PromptChar":        // maskedit
                    validProperty = false;
                    break;

                // begin common properties

                case "Alignment":
                    //              0 - left
                    //              1 - right
                    //              2 - center
                    targetProperty.Name = "TextAlign";
                    targetProperty.Value = "System.Drawing.ContentAlignment.";
                    switch (sourceProperty.Value)
                    {
                        case "0":
                            targetProperty.Value = targetProperty.Value + "TopLeft";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "TopRight";
                            break;

                        case "2":
                        default:
                            targetProperty.Value = targetProperty.Value + "TopCenter";
                            break;
                    }
                    break;

                case "BackColor":
                case "ForeColor":
                    if (type != "ImageList")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = GetColor(sourceProperty.Value);
                    }
                    else
                    {
                        validProperty = false;
                    }
                    break;

                case "BorderStyle":
                    if (type == "form")
                    {
                        targetProperty.Name = "FormBorderStyle";
                        // 0 - none
                        // 1 - fixed single
                        // 2 - sizable
                        // 3 - fixed dialog
                        // 4 - fixed toolwindow
                        // 5 - sizable toolwindow

                        // FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;

                        targetProperty.Value = "System.Windows.Forms.FormBorderStyle.";
                        switch (sourceProperty.Value)
                        {
                            case "0":
                                targetProperty.Value = targetProperty.Value + "None";
                                break;

                            default:
                            case "1":
                                targetProperty.Value = targetProperty.Value + "FixedSingle";
                                break;

                            case "2":
                                targetProperty.Value = targetProperty.Value + "Sizable";
                                break;

                            case "3":
                                targetProperty.Value = targetProperty.Value + "FixedDialog";
                                break;

                            case "4":
                                targetProperty.Value = targetProperty.Value + "FixedToolWindow";
                                break;

                            case "5":
                                targetProperty.Value = targetProperty.Value + "SizableToolWindow";
                                break;
                        }
                    }
                    else
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = "System.Windows.Forms.BorderStyle.";
                        switch (sourceProperty.Value)
                        {
                            case "0":
                                targetProperty.Value = targetProperty.Value + "None";
                                break;

                            case "1":
                                targetProperty.Value = targetProperty.Value + "FixedSingle";
                                break;

                            case "2":
                            default:
                                targetProperty.Value = targetProperty.Value + "Fixed3D";
                                break;
                        }
                    }
                    break;

                case "Caption":
                case "Text":
                    targetProperty.Name = "Text";
                    targetProperty.Value = sourceProperty.Value;
                    break;

                // this.cmdExit.Size = new System.Drawing.Size(80, 40);
                case "Height":
                    targetProperty.Name = "Size";
                    targetProperty.Value = "new System.Drawing.Size(" + GetSize("Height", "Width", sourcePropertyList) + ")";
                    break;

                // this.cmdExit.Location = new System.Drawing.Point(616, 520);
                case "Left":
                    if ((type != "ImageList") && (type != "Timer"))
                    {
                        targetProperty.Name = "Location";
                        targetProperty.Value = "new System.Drawing.Point(" + GetLocation(sourcePropertyList) + ")";
                    }
                    else
                    {
                        validProperty = false;
                    }
                    break;

                case "Top":
                case "Width":
                    // nothing, already processed by Height, Left
                    validProperty = false;
                    break;

                case "Enabled":
                case "Locked":
                case "TabStop":
                case "Visible":
                case "UseMnemonic":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "WordWrap":
                    if (type == "Text")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = GetBool(sourceProperty.Value);
                    }
                    else
                    {
                        validProperty = false;
                    }
                    break;

                case "Font":
                    ConvertFont(sourceProperty, targetProperty);
                    break;
                // end common properties

                case "MaxLength":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = sourceProperty.Value;
                    break;

                // PasswordChar
                case "PasswordChar":
                    targetProperty.Name = sourceProperty.Name;
                    // PasswordChar = '*';
                    targetProperty.Value = "'" + sourceProperty.Value.Substring(1, 1) + "'";
                    break;

                // Value
                case "Value":
                    switch (type)
                    {
                        case "RadioButton":
                            // .Checked = true;
                            targetProperty.Name = "Checked";
                            targetProperty.Value = GetBool(sourceProperty.Value);
                            break;

                        case "CheckBox":
                            //.CheckState = System.Windows.Forms.CheckState.Checked;
                            targetProperty.Name = "CheckState";
                            targetProperty.Value = "System.Windows.Forms.CheckState.";
                            // 0 - Unchecked
                            // 1 - checked
                            // 2 - grayed
                            switch (sourceProperty.Value)
                            {
                                default:
                                case "0":
                                    targetProperty.Value = targetProperty.Value + "Unchecked";
                                    break;

                                case "1":
                                    targetProperty.Value = targetProperty.Value + "Checked";
                                    break;

                                case "2":
                                    targetProperty.Value = targetProperty.Value + "Indeterminate";
                                    break;
                            }
                            break;

                        default:
                            targetProperty.Value = targetProperty.Value + "Both";
                            break;
                    }
                    break;

                // timer
                case "Interval":
                    targetProperty.Name = "Interval";
                    targetProperty.Value = sourceProperty.Value;
                    break;

                // this.cmdExit.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
                case "Cancel":
                    if (int.Parse(sourceProperty.Value) != 0)
                    {
                        targetProperty.Name = "DialogResult";
                        targetProperty.Value = "System.Windows.Forms.DialogResult.Cancel";
                    }
                    break;

                case "Default":
                    if (int.Parse(sourceProperty.Value) != 0)
                    {
                        targetProperty.Name = "DialogResult";
                        targetProperty.Value = "System.Windows.Forms.DialogResult.OK";
                    }
                    break;

                //                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
                //                this.ClientSize = new System.Drawing.Size(704, 565);
                //                this.MinimumSize = new System.Drawing.Size(712, 592);

                // direct value
                case "TabIndex":
                case "Tag":
                    // except MenuItem
                    if (type != "MenuItem")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = sourceProperty.Value;
                    }
                    else
                    {
                        validProperty = false;
                    }
                    break;

                // -1 converted to true
                // 0 to false
                case "AutoSize":
                    // only for Label
                    if (type == "Label")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = GetBool(sourceProperty.Value);
                    }
                    else
                    {
                        validProperty = false;
                    }
                    break;

                case "Icon":
                    // "Form1.frx":0000;
                    // exist file ?

                    //          System.Drawing.Bitmap pic = null;
                    //          GetFRXImage(@"C:\temp\test\form1.frx", 0x13960, pic );

                    if (type == "form")
                    {
                        //.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
                        targetProperty.Name = "Icon";
                        targetProperty.Value = sourceProperty.Value;
                    }
                    else
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
                        targetProperty.Name = "Image";
                        targetProperty.Value = sourceProperty.Value;
                    }
                    break;

                case "Picture":
                    // = "Form1.frx":13960;
                    if (type == "form")
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("$this.BackgroundImage")));
                        targetProperty.Name = "BackgroundImage";
                        targetProperty.Value = sourceProperty.Value;
                    }
                    else
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
                        targetProperty.Name = "Image";
                        targetProperty.Value = sourceProperty.Value;
                    }
                    break;

                case "ScrollBars":
                    // ScrollBars = System.Windows.Forms.ScrollBars.Both;
                    targetProperty.Name = sourceProperty.Name;

                    if (type == "RichTextBox")
                    {
                        targetProperty.Value = "System.Windows.Forms.RichTextBoxScrollBars.";
                    }
                    else
                    {
                        targetProperty.Value = "System.Windows.Forms.ScrollBars.";
                    }
                    switch (sourceProperty.Value)
                    {
                        default:
                        case "0":
                            targetProperty.Value = targetProperty.Value + "None";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "Horizontal";
                            break;

                        case "2":
                            targetProperty.Value = targetProperty.Value + "Vertical";
                            break;

                        case "3":
                            targetProperty.Value = targetProperty.Value + "Both";
                            break;
                    }
                    break;

                // SS tab
                case "TabOrientation":
                    targetProperty.Name = "Alignment";
                    targetProperty.Value = "System.Windows.Forms.TabAlignment.";
                    switch (sourceProperty.Value)
                    {
                        default:
                        case "0":
                            targetProperty.Value = targetProperty.Value + "Top";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "Bottom";
                            break;

                        case "2":
                            targetProperty.Value = targetProperty.Value + "Left";
                            break;

                        case "3":
                            targetProperty.Value = targetProperty.Value + "Right";
                            break;
                    }
                    break;

                // begin Listview

                // unsupported properties
                case "_ExtentX":
                case "_ExtentY":
                case "_Version":
                case "OLEDropMode":
                    validProperty = false;
                    break;

                // this.listView.View = System.Windows.Forms.View.List;
                case "View":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = "System.Windows.Forms.View.";
                    switch (sourceProperty.Value)
                    {
                        case "0":
                            targetProperty.Value = targetProperty.Value + "Details";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "LargeIcon";
                            break;

                        case "2":
                            targetProperty.Value = targetProperty.Value + "SmallIcon";
                            break;

                        case "3":
                        default:
                            targetProperty.Value = targetProperty.Value + "List";
                            break;
                    }
                    break;

                case "LabelEdit":
                case "LabelWrap":
                case "MultiSelect":
                case "HideSelection":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                // end List view

                // VB6 form unsupported properties
                case "MDIChild":
                case "WhatsThisButton":
                case "NegotiateMenus":
                case "HelpContextID":
                case "LinkTopic":
                case "PaletteMode":
                case "ClipControls":
                case "LockControls":
                case "FillStyle":
                    validProperty = false;
                    break;

                // supported properties

                case "ControlBox":
                case "KeyPreview":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "ClientHeight":
                    targetProperty.Name = "ClientSize";
                    targetProperty.Value = "new System.Drawing.Size(" + GetSize("ClientHeight", "ClientWidth", sourcePropertyList) + ")";
                    break;

                case "ClientWidth":
                    // nothing, already processed by Height, Left
                    validProperty = false;
                    break;

                case "ClientLeft":
                case "ClientTop":
                    validProperty = false;
                    break;

                case "MaxButton":
                    targetProperty.Name = "MaximizeBox";
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "MinButton":
                    targetProperty.Name = "MinimizeBox";
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "WhatsThisHelp":
                    targetProperty.Name = "HelpButton";
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "ShowInTaskbar":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                case "WindowList":
                    targetProperty.Name = "MdiList";
                    targetProperty.Value = GetBool(sourceProperty.Value);
                    break;

                // this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                // 0 - normal
                // 1 - minimized
                // 2 - maximized
                case "WindowState":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = "System.Windows.Forms.FormWindowState.";
                    switch (sourceProperty.Value)
                    {
                        case "0":
                        default:
                            targetProperty.Value = targetProperty.Value + "Normal";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "Minimized";
                            break;

                        case "2":
                            targetProperty.Value = targetProperty.Value + "Maximized";
                            break;
                    }
                    break;

                case "StartUpPosition":
                    // 0 - manual
                    // 1 - center owner
                    // 2 - center screen
                    // 3 - windows default
                    targetProperty.Name = "StartPosition";
                    targetProperty.Value = "System.Windows.Forms.FormStartPosition.";
                    switch (sourceProperty.Value)
                    {
                        case "0":
                            targetProperty.Value = targetProperty.Value + "Manual";
                            break;

                        case "1":
                            targetProperty.Value = targetProperty.Value + "CenterParent";
                            break;

                        case "2":
                            targetProperty.Value = targetProperty.Value + "CenterScreen";
                            break;

                        case "3":
                        default:
                            targetProperty.Value = targetProperty.Value + "WindowsDefaultLocation";
                            break;
                    }
                    break;

                default:
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = sourceProperty.Value;
                    targetProperty.Valid = false;
                    break;
            }
            return validProperty;
        }

        public static bool ParseVariable(Variable sourceVariable, Variable targetVariable)
        {
            targetVariable.Scope = sourceVariable.Scope;
            targetVariable.Name = sourceVariable.Name;
            targetVariable.Type = VariableTypeConvert(sourceVariable.Type);

            return true;
        }

        public static bool ParseVariables(ArrayList sourceVariableList, ArrayList targetVariableList)
        {
            // each property
            foreach (Variable sourceVariable in sourceVariableList)
            {
                var targetVariable = new Variable();
                if (ParseVariable(sourceVariable, targetVariable))
                {
                    targetVariableList.Add(targetVariable);
                }
            }
            return true;
        }

        private static void ControlListLoad()
        {
            _mControlList = new Hashtable();
            var doc = new XmlDocument();
            XmlNode node;
            ControlListItem oItem;

            // get current directory
            string[] commandLineArgs;
            commandLineArgs = Environment.GetCommandLineArgs();
            // index 0 contain path and name of exe file
            var binPath = Path.GetDirectoryName(commandLineArgs[0].ToLower());
            var fileName = binPath + @"\vb2c.xml";

            doc.Load(fileName);
            // Select the node given
            node = doc.DocumentElement.SelectSingleNode("/configuration/ControlList");
            // exit with an empty collection if nothing here
            if (node == null) { return; }
            // exit with an empty colection if the node has no children
            if (node.HasChildNodes == false) { return; }
            // get the nodelist of all children
            var nodeList = node.ChildNodes;

            foreach (XmlElement element in nodeList)
            {
                oItem = new ControlListItem();
                oItem.Vb6Name = string.Empty;
                oItem.CsharpName = string.Empty;
                oItem.Unsupported = false;
                oItem.InvisibleAtRuntime = false;
                foreach (XmlElement childElement in element)
                {
                    switch (childElement.Name)
                    {
                        case "VB6":
                            // compare in uppercase
                            oItem.Vb6Name = childElement.InnerText.ToUpper();
                            break;

                        case "Csharp":
                            oItem.CsharpName = childElement.InnerText;
                            break;

                        case "Unsupported":
                            oItem.Unsupported = bool.Parse(childElement.InnerText);
                            break;

                        case "InvisibleAtRuntime":
                            oItem.InvisibleAtRuntime = bool.Parse(childElement.InnerText);
                            break;
                    }
                }
                _mControlList.Add(oItem.Vb6Name, oItem);
            }

            //      private string getKeyValue(string aSection, string aKey, string aDefaultValue)
            //      {
            //        XmlNode node;
            //        node = (Doc.DocumentElement).SelectSingleNode("/configuration/" + aSection + "/" + aKey);
            //        if (node == null) {return aDefaultValue;}
            //        return node.InnerText;
            //      }
        }

        private static void ConvertFont(ControlProperty sourceProperty, ControlProperty targetProperty)
        {
            var fontName = string.Empty;
            var fontSize = 0;
            var fontCharSet = 0;
            var fontBold = false;
            var fontUnderline = false;
            var fontItalic = false;
            var fontStrikethrough = false;
            var temp = string.Empty;
            //      BeginProperty Font
            //         Name            =   "Arial"
            //         Size            =   8.25
            //         Charset         =   238
            //         Weight          =   400
            //         Underline       =   0   'False
            //         Italic          =   0   'False
            //         Strikethrough   =   0   'False
            //      EndProperty

            foreach (ControlProperty oProperty in sourceProperty.PropertyList)
            {
                switch (oProperty.Name)
                {
                    case "Name":
                        fontName = oProperty.Value;
                        break;

                    case "Size":
                        fontSize = GetFontSizeInt(oProperty.Value);
                        break;

                    case "Weight":
                        //        If tLogFont.lfWeight >= FW_BOLD Then
                        //          bFontBold = True
                        //        Else
                        //          bFontBold = False
                        //        End If
                        // FW_BOLD = 700
                        fontBold = (int.Parse(oProperty.Value) >= 700);
                        break;

                    case "Charset":
                        fontCharSet = int.Parse(oProperty.Value);
                        break;

                    case "Underline":
                        fontUnderline = (int.Parse(oProperty.Value) != 0);
                        break;

                    case "Italic":
                        fontItalic = (int.Parse(oProperty.Value) != 0);
                        break;

                    case "Strikethrough":
                        fontStrikethrough = (int.Parse(oProperty.Value) != 0);
                        break;
                }
            }

            //      this.cmdExit.Font = new System.Drawing.Font("Tahoma", 12F,
            //        (System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline
            //        | System.Drawing.FontStyle.Strikeout), System.Drawing.GraphicsUnit.Point,
            //        ((System.Byte)(0)));

            // this.cmdExit.Font = new System.Drawing.Font("Tahoma", 12F,
            // System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point,
            // ((System.Byte)(238)));

            targetProperty.Name = "Font";
            targetProperty.Value = "new System.Drawing.Font(" + fontName + ",";
            targetProperty.Value = targetProperty.Value + fontSize.ToString() + "F,";

            temp = string.Empty;
            if (fontBold)
            {
                temp = "System.Drawing.FontStyle.Bold";
            }
            if (fontItalic)
            {
                if (temp != string.Empty) { temp = temp + " | "; }
                temp = temp + "System.Drawing.FontStyle.Italic";
            }
            if (fontUnderline)
            {
                if (temp != string.Empty) { temp = temp + " | "; }
                temp = temp + "System.Drawing.FontStyle.Underline";
            }
            if (fontStrikethrough)
            {
                if (temp != string.Empty) { temp = temp + " | "; }
                temp = temp + "System.Drawing.FontStyle.Strikeout";
            }
            if (temp == string.Empty)
            {
                targetProperty.Value = targetProperty.Value + " System.Drawing.FontStyle.Regular,";
            }
            else
            {
                targetProperty.Value = targetProperty.Value + " ( " + temp + " ),";
            }
            targetProperty.Value = targetProperty.Value + " System.Drawing.GraphicsUnit.Point, ";
            targetProperty.Value = targetProperty.Value + "((System.Byte)(" + fontCharSet.ToString() + ")));";
        }

        private static string GetBool(string value)
        {
            if (int.Parse(value) == 0)
            {
                return "false";
            }
            else
            {
                return "true";
            }
        }

        private static string GetColor(string value)
        {
            Color color;

            if (value.Length < 3)
            {
                var colorValue = "0x" + value;
                color = ColorTranslator.FromWin32(Convert.ToInt32(colorValue, 16));
            }
            else if (value.StartsWith("&"))
            {
                value = value
                    .Replace("&H", "")
                    .Replace("&", "")
                    ;
                color = ColorTranslator.FromHtml("#" + value);
            }
            else
            {
                color = Color.FromArgb(Convert.ToInt32(value));

                //Color = System.Drawing.ColorTranslator.FromWin32(System.Convert.ToInt32(Value, 16));
            }

            if (!color.IsSystemColor)
            {
                if (color.IsNamedColor)
                {
                    // System.Drawing.Color.Yellow;
                    return "System.Drawing.Color." + color.Name;
                }
                else
                {
                    return "System.Drawing.Color.FromArgb(" + color.ToArgb() + ")";
                }
            }
            else
            {
                return "System.Drawing.SystemColors." + color.Name;
            }
        }

        // return control name
        private static string GetControlIndexName(string tabName)
        {
            //  this.SSTab1.(Tab(1).Control(4) = "Option1(0)";
            var start = 0;
            var end = 0;

            start = tabName.IndexOf("(", 0);
            if (start > -1)
            {
                end = tabName.IndexOf(")", 0);
                return tabName.Substring(0, start) + tabName.Substring(start + 1, end - start - 1);
            }
            else
            {
                return tabName;
            }
        }

        private static int GetFontSizeInt(string value)
        {
            var position = 0;

            position = value.IndexOf(",", 0);
            if (position > -1)
            {
                return int.Parse(value.Substring(0, position));
            }

            position = value.IndexOf(".", 0);
            if (position > 0)
            {
                return int.Parse(value.Substring(0, position));
            }
            return int.Parse(value);
        }

        private static string GetLocation(ArrayList propertyList)
        {
            var left = 0;
            var top = 0;

            // each property
            foreach (ControlProperty oProperty in propertyList)
            {
                if (oProperty.Name == "Left")
                {
                    left = int.Parse(oProperty.Value);
                    if (left < 0)
                    {
                        left = 75000 + left;
                    }
                    left = left / 15;
                }
                if (oProperty.Name == "Top")
                {
                    top = int.Parse(oProperty.Value) / 15;
                }
            }
            // 616, 520
            return left.ToString() + ", " + top.ToString();
        }

        private static string GetSize(string height, string width, ArrayList propertyList)
        {
            var heightValue = 0;
            var widthValue = 0;

            // each property
            foreach (ControlProperty oProperty in propertyList)
            {
                if (oProperty.Name == height)
                {
                    heightValue = int.Parse(oProperty.Value) / 15;
                }
                if (oProperty.Name == width)
                {
                    widthValue = int.Parse(oProperty.Value) / 15;
                }
            }
            // 0, 120
            return widthValue.ToString() + ", " + heightValue.ToString();
        }

        private static string VariableTypeConvert(string sourceType)
        {
            string targetType;

            switch (sourceType)
            {
                case "Long":
                    targetType = "int";
                    break;

                case "Integer":
                    targetType = "short";
                    break;

                case "Byte":
                    targetType = "byte";
                    break;

                case "String":
                    targetType = "string";
                    break;

                case "Boolean":
                    targetType = "bool";
                    break;

                case "Currency":
                    targetType = "decimal";
                    break;

                case "Single":
                    targetType = "float";
                    break;

                case "Double":
                    targetType = "double";
                    break;

                case "ADODB.Recordset":
                case "DAO.Recordset":
                case "Recordset":
                    targetType = "DataReader";
                    break;

                default:
                    targetType = sourceType;
                    break;
            }
            return targetType;
        }
    }
}