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

        public static void GetFrxImage(string ImageFile, int ImageOffset, out byte[] ImageString)
        {
            byte[] Header;
            var BytesToRead = 0;

            // open file
            var Stream = new FileStream(ImageFile, FileMode.Open, FileAccess.Read);
            var Reader = new BinaryReader(Stream);
            // Start from offset
            Reader.BaseStream.Seek(ImageOffset, SeekOrigin.Begin);
            // Get the four byte header
            Header = new byte[4];
            Header = Reader.ReadBytes(4);
            // Convert This Header Into The Number Of Bytes
            // To Read For This Image
            BytesToRead = Header[0];
            BytesToRead = BytesToRead + (Header[1] * 0x100);
            BytesToRead = BytesToRead + (Header[2] * 0x10000);
            BytesToRead = BytesToRead + (Header[3] * 0x1000000);
            // Get image information
            ImageString = new byte[BytesToRead];
            ImageString = Reader.ReadBytes(BytesToRead);

            //      Stream = new FileStream( @"C:\temp\test\Ba.bmp", FileMode.CreateNew, FileAccess.Write );
            //      BinaryWriter Writer = new BinaryWriter( Stream );
            //      Writer.Write(ImageString, 8, ImageString.GetLength(0) - 8);
            //      Stream.Close();
            //      Writer.Close();

            //  FileStream inFile = new FileStream(@"C:\WINdows\Blue Lace 16.bmp", FileMode.Open, FileAccess.Read);
            //  ReturnImage = Image.FromStream(inFile, false);

            Stream.Close();
            Reader.Close();
        }

        public static bool ParseClassProperties(Module SourceModule, Module TargetModule)
        {
            foreach (Property SourceProperty in SourceModule.PropertyList)
            {
                var TargetProperty = new Property
                {
                    Name = SourceProperty.Name, 
                    Comment = SourceProperty.Comment,
                    Scope = SourceProperty.Scope,
                    Type = VariableTypeConvert(SourceProperty.Type),
                };

                switch (SourceProperty.Direction)
                {
                    case "Get":
                        TargetProperty.Direction = "get";
                        break;

                    case "Set":
                    case "Let":
                        TargetProperty.Direction = "set";
                        break;
                }
                // lines
                foreach (string Line in SourceProperty.LineList)
                    if (Line.Trim() != string.Empty)
                        TargetProperty.LineList.Add(Line);

                TargetModule.PropertyList.Add(TargetProperty);
            }
            return true;
        }

        public static bool ParseControlProperties(Module oModule, Control oControl,
                                                ArrayList SourcePropertyList,
                                                ArrayList TargetPropertyList)
        {
            ControlProperty TargetProperty = null;

            // each property
            foreach (ControlProperty SourceProperty in SourcePropertyList)
            {
                if (SourceProperty.Name == "Index")
                {
                    // Index           =   3
                    oControl.Name = oControl.Name + SourceProperty.Value;
                }
                else
                {
                    TargetProperty = new ControlProperty();
                    if (ParseProperties(oControl.Type, SourceProperty, TargetProperty, SourcePropertyList))
                    {
                        if (TargetProperty.Name == "Image")
                        {
                            oModule.ImagesUsed = true;
                        }
                        TargetPropertyList.Add(TargetProperty);
                    }
                }
            }
            return true;
        }

        public static bool ParseControls(Module oModule, ArrayList SourceControlList, ArrayList TargetControlList)
        {
            var Type = string.Empty;

            foreach (Control oSourceControl in SourceControlList)
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
                        Type = "Unsuported";
                        oTargetControl.Valid = false;
                    }
                    else
                    {
                        Type = oItem.CsharpName;
                        if (Type == "MenuItem")
                        {
                            oModule.MenuUsed = true;
                        }
                    }
                    oTargetControl.InvisibleAtRuntime = oItem.InvisibleAtRuntime;
                }
                else
                {
                    Type = oSourceControl.Type;
                }

                oTargetControl.Type = Type;
                ParseControlProperties(oModule, oTargetControl, oSourceControl.PropertyList, oTargetControl.PropertyList);

                TargetControlList.Add(oTargetControl);
            }
            return true;
        }

        public static bool ParseEnums(Module SourceModule, Module TargetModule)
        {
            foreach (Enum SourceEnum in SourceModule.EnumList)
            {
                TargetModule.EnumList.Add(SourceEnum);
            }
            return true;
        }

        public static bool ParseModule(Module SourceModule, Module targetModule)
        {
            ControlListLoad();

            // module name
            targetModule.Name = SourceModule.Name;
            // file name
            targetModule.FileName = Path.GetFileNameWithoutExtension(SourceModule.FileName) + ".cs";
            // type
            targetModule.Type = SourceModule.Type;
            // version
            targetModule.Version = SourceModule.Version;
            // process own properties - forms
            ParseModuleProperties(targetModule, SourceModule.FormPropertyList, targetModule.FormPropertyList);
            // process controls - form
            ParseControls(targetModule, SourceModule.ControlList, targetModule.ControlList);

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

            var TempControlList = new ArrayList();
            var TabControlIndex = 0;

            // check for TabDlg.SSTab
            foreach (Control oTargetControl in targetModule.ControlList)
            {
                if ((oTargetControl.Type == "TabControl") && (oTargetControl.Valid))
                {
                    // for each source table is necessary
                    //          this.tabControl1 = new System.Windows.Forms.TabControl();
                    //          this.tabPage1 = new System.Windows.Forms.TabPage();

                    var Index = 0;
                    Control oTabPage = null;
                    // each property
                    foreach (ControlProperty oTargetProperty in oTargetControl.PropertyList)
                    {
                        // TabCaption = create new tab
                        //      this.SSTab1.(TabCaption(0)) = "Tab 0";

                        Console.WriteLine(oTargetProperty.Name);

                        if (oTargetProperty.Name.IndexOf("TabCaption(" + Index.ToString() + ")", 0) > -1)
                        {
                            // new tab
                            oTabPage = new Control();
                            oTabPage.Type = "TabPage";
                            oTabPage.Name = "tabPage" + Index.ToString();
                            oTabPage.Owner = oTargetControl.Name;
                            oTabPage.Container = true;
                            oTabPage.Valid = true;
                            oTabPage.InvisibleAtRuntime = false;

                            // add some necessary properties
                            var TargetProperty = new ControlProperty();
                            TargetProperty.Name = "Location";
                            TargetProperty.Value = "new System.Drawing.Point(4, 22)";
                            TargetProperty.Valid = true;
                            oTabPage.PropertyList.Add(TargetProperty);

                            TargetProperty = new ControlProperty();
                            TargetProperty.Name = "Size";
                            TargetProperty.Value = "new System.Drawing.Size(477, 374)";
                            TargetProperty.Valid = true;
                            oTabPage.PropertyList.Add(TargetProperty);

                            TargetProperty = new ControlProperty();
                            TargetProperty.Name = "Text";
                            TargetProperty.Value = oTargetProperty.Value;
                            TargetProperty.Valid = true;
                            oTabPage.PropertyList.Add(TargetProperty);

                            TargetProperty = new ControlProperty();
                            TargetProperty.Name = "TabIndex";
                            TargetProperty.Value = Index.ToString();
                            TargetProperty.Valid = true;
                            oTabPage.PropertyList.Add(TargetProperty);

                            TempControlList.Add(oTabPage);
                            Index++;
                        }

                        // Control = change owner of control to current tab
                        //      this.SSTab1.(Tab(0).Control(0) = "ImageControl";
                        if (oTargetProperty.Name.IndexOf(".Control(", 0) > -1)
                        {
                            if (oTargetProperty.Name.IndexOf("Enable", 0) == -1)
                            {
                                var TabName = oTargetProperty.Value.Substring(1, oTargetProperty.Value.Length - 2);
                                TabName = GetControlIndexName(TabName);
                                // search for "oTargetProperty.Value" control
                                // and replace owner of this control to current tab
                                foreach (Control oNewOwner in targetModule.ControlList)
                                {
                                    if ((oNewOwner.Name == TabName) && (!oNewOwner.InvisibleAtRuntime))
                                    {
                                        oNewOwner.Owner = oTabPage.Name;
                                    }
                                }
                            }
                        }
                    }
                }
                TabControlIndex++;
            }

            if (TempControlList.Count > 0)
            {
                // right order of tabs
                var Position = 0;
                foreach (Control oControl in TempControlList)
                {
                    targetModule.ControlList.Insert(TabControlIndex + Position, oControl);
                    Position++;
                }
            }

            // process enums
            ParseEnums(SourceModule, targetModule);
            // process variables
            ParseVariables(SourceModule.VariableList, targetModule.VariableList);
            // process properties
            ParseClassProperties(SourceModule, targetModule);
            // process procedures
            ParseProcedures(SourceModule, targetModule);

            return true;
        }

        public static bool ParseModuleProperties(Module oModule,
                                              ArrayList SourcePropertyList,
                                              ArrayList TargetPropertyList)
        {
            // each property
            foreach (ControlProperty SourceProperty in SourcePropertyList)
            {
                var TargetProperty = new ControlProperty();
                if (ParseProperties(oModule.Type, SourceProperty, TargetProperty, SourcePropertyList))
                {
                    if (TargetProperty.Name == "BackgroundImage" || TargetProperty.Name == "Icon")
                    {
                        oModule.ImagesUsed = true;
                    }
                    TargetPropertyList.Add(TargetProperty);
                }
            }
            return true;
        }

        public static bool ParseProcedures(Module SourceModule, Module TargetModule)
        {
            const string indent6 = "      ";

            foreach (Procedure SourceProcedure in SourceModule.ProcedureList)
            {
                var TargetProcedure = new Procedure
                {
                    Name = SourceProcedure.Name,
                    Scope = SourceProcedure.Scope,
                    Comment = SourceProcedure.Comment,
                    Type = SourceProcedure.Type,
                    ReturnType = VariableTypeConvert(SourceProcedure.ReturnType),
                    ParameterList = SourceProcedure.ParameterList
                };
                
                // lines
                foreach (string Line in SourceProcedure.LineList)
                {
                    var tempSource = Line.Trim();
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
                        if (Line.Contains("If ") 
                            && Line.Contains("Then") 
                            && Line.TokensAfter(1, "Then").Trim().Length > 0)
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
                            TargetProcedure.BottomLineList.Add(@"
        catch(Exception e) 
        { 
            /* ON ERROR RESUME NEXT (ish) */ 
        }
");
                            tempSource = translatedLine = "try\r\n{\r\n";
                        }
                        else
                        {
                            TargetProcedure.LineList.Add(translatedLine == string.Empty ? tempSource : translatedLine);
                        }
                    }
                    else
                    {
                        TargetProcedure.LineList.Add(string.Empty);
                    }
                }

                TargetModule.ProcedureList.Add(TargetProcedure);
            }
            return true;
        }

        public static bool ParseProperties(string Type,
                                        ControlProperty SourceProperty,
                                        ControlProperty TargetProperty,
                                        ArrayList SourcePropertyList)
        {
            var ValidProperty = true;
            TargetProperty.Valid = true;

            switch (SourceProperty.Name)
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
                    ValidProperty = false;
                    break;

                // begin common properties

                case "Alignment":
                    //              0 - left
                    //              1 - right
                    //              2 - center
                    TargetProperty.Name = "TextAlign";
                    TargetProperty.Value = "System.Drawing.ContentAlignment.";
                    switch (SourceProperty.Value)
                    {
                        case "0":
                            TargetProperty.Value = TargetProperty.Value + "TopLeft";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "TopRight";
                            break;

                        case "2":
                        default:
                            TargetProperty.Value = TargetProperty.Value + "TopCenter";
                            break;
                    }
                    break;

                case "BackColor":
                case "ForeColor":
                    if (Type != "ImageList")
                    {
                        TargetProperty.Name = SourceProperty.Name;
                        TargetProperty.Value = GetColor(SourceProperty.Value);
                    }
                    else
                    {
                        ValidProperty = false;
                    }
                    break;

                case "BorderStyle":
                    if (Type == "form")
                    {
                        TargetProperty.Name = "FormBorderStyle";
                        // 0 - none
                        // 1 - fixed single
                        // 2 - sizable
                        // 3 - fixed dialog
                        // 4 - fixed toolwindow
                        // 5 - sizable toolwindow

                        // FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;

                        TargetProperty.Value = "System.Windows.Forms.FormBorderStyle.";
                        switch (SourceProperty.Value)
                        {
                            case "0":
                                TargetProperty.Value = TargetProperty.Value + "None";
                                break;

                            default:
                            case "1":
                                TargetProperty.Value = TargetProperty.Value + "FixedSingle";
                                break;

                            case "2":
                                TargetProperty.Value = TargetProperty.Value + "Sizable";
                                break;

                            case "3":
                                TargetProperty.Value = TargetProperty.Value + "FixedDialog";
                                break;

                            case "4":
                                TargetProperty.Value = TargetProperty.Value + "FixedToolWindow";
                                break;

                            case "5":
                                TargetProperty.Value = TargetProperty.Value + "SizableToolWindow";
                                break;
                        }
                    }
                    else
                    {
                        TargetProperty.Name = SourceProperty.Name;
                        TargetProperty.Value = "System.Windows.Forms.BorderStyle.";
                        switch (SourceProperty.Value)
                        {
                            case "0":
                                TargetProperty.Value = TargetProperty.Value + "None";
                                break;

                            case "1":
                                TargetProperty.Value = TargetProperty.Value + "FixedSingle";
                                break;

                            case "2":
                            default:
                                TargetProperty.Value = TargetProperty.Value + "Fixed3D";
                                break;
                        }
                    }
                    break;

                case "Caption":
                case "Text":
                    TargetProperty.Name = "Text";
                    TargetProperty.Value = SourceProperty.Value;
                    break;

                // this.cmdExit.Size = new System.Drawing.Size(80, 40);
                case "Height":
                    TargetProperty.Name = "Size";
                    TargetProperty.Value = "new System.Drawing.Size(" + GetSize("Height", "Width", SourcePropertyList) + ")";
                    break;

                // this.cmdExit.Location = new System.Drawing.Point(616, 520);
                case "Left":
                    if ((Type != "ImageList") && (Type != "Timer"))
                    {
                        TargetProperty.Name = "Location";
                        TargetProperty.Value = "new System.Drawing.Point(" + GetLocation(SourcePropertyList) + ")";
                    }
                    else
                    {
                        ValidProperty = false;
                    }
                    break;

                case "Top":
                case "Width":
                    // nothing, already processed by Height, Left
                    ValidProperty = false;
                    break;

                case "Enabled":
                case "Locked":
                case "TabStop":
                case "Visible":
                case "UseMnemonic":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "WordWrap":
                    if (Type == "Text")
                    {
                        TargetProperty.Name = SourceProperty.Name;
                        TargetProperty.Value = GetBool(SourceProperty.Value);
                    }
                    else
                    {
                        ValidProperty = false;
                    }
                    break;

                case "Font":
                    ConvertFont(SourceProperty, TargetProperty);
                    break;
                // end common properties

                case "MaxLength":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = SourceProperty.Value;
                    break;

                // PasswordChar
                case "PasswordChar":
                    TargetProperty.Name = SourceProperty.Name;
                    // PasswordChar = '*';
                    TargetProperty.Value = "'" + SourceProperty.Value.Substring(1, 1) + "'";
                    break;

                // Value
                case "Value":
                    switch (Type)
                    {
                        case "RadioButton":
                            // .Checked = true;
                            TargetProperty.Name = "Checked";
                            TargetProperty.Value = GetBool(SourceProperty.Value);
                            break;

                        case "CheckBox":
                            //.CheckState = System.Windows.Forms.CheckState.Checked;
                            TargetProperty.Name = "CheckState";
                            TargetProperty.Value = "System.Windows.Forms.CheckState.";
                            // 0 - Unchecked
                            // 1 - checked
                            // 2 - grayed
                            switch (SourceProperty.Value)
                            {
                                default:
                                case "0":
                                    TargetProperty.Value = TargetProperty.Value + "Unchecked";
                                    break;

                                case "1":
                                    TargetProperty.Value = TargetProperty.Value + "Checked";
                                    break;

                                case "2":
                                    TargetProperty.Value = TargetProperty.Value + "Indeterminate";
                                    break;
                            }
                            break;

                        default:
                            TargetProperty.Value = TargetProperty.Value + "Both";
                            break;
                    }
                    break;

                // timer
                case "Interval":
                    TargetProperty.Name = "Interval";
                    TargetProperty.Value = SourceProperty.Value;
                    break;

                // this.cmdExit.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
                case "Cancel":
                    if (int.Parse(SourceProperty.Value) != 0)
                    {
                        TargetProperty.Name = "DialogResult";
                        TargetProperty.Value = "System.Windows.Forms.DialogResult.Cancel";
                    }
                    break;

                case "Default":
                    if (int.Parse(SourceProperty.Value) != 0)
                    {
                        TargetProperty.Name = "DialogResult";
                        TargetProperty.Value = "System.Windows.Forms.DialogResult.OK";
                    }
                    break;

                //                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
                //                this.ClientSize = new System.Drawing.Size(704, 565);
                //                this.MinimumSize = new System.Drawing.Size(712, 592);

                // direct value
                case "TabIndex":
                case "Tag":
                    // except MenuItem
                    if (Type != "MenuItem")
                    {
                        TargetProperty.Name = SourceProperty.Name;
                        TargetProperty.Value = SourceProperty.Value;
                    }
                    else
                    {
                        ValidProperty = false;
                    }
                    break;

                // -1 converted to true
                // 0 to false
                case "AutoSize":
                    // only for Label
                    if (Type == "Label")
                    {
                        TargetProperty.Name = SourceProperty.Name;
                        TargetProperty.Value = GetBool(SourceProperty.Value);
                    }
                    else
                    {
                        ValidProperty = false;
                    }
                    break;

                case "Icon":
                    // "Form1.frx":0000;
                    // exist file ?

                    //          System.Drawing.Bitmap pic = null;
                    //          GetFRXImage(@"C:\temp\test\form1.frx", 0x13960, pic );

                    if (Type == "form")
                    {
                        //.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
                        TargetProperty.Name = "Icon";
                        TargetProperty.Value = SourceProperty.Value;
                    }
                    else
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
                        TargetProperty.Name = "Image";
                        TargetProperty.Value = SourceProperty.Value;
                    }
                    break;

                case "Picture":
                    // = "Form1.frx":13960;
                    if (Type == "form")
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("$this.BackgroundImage")));
                        TargetProperty.Name = "BackgroundImage";
                        TargetProperty.Value = SourceProperty.Value;
                    }
                    else
                    {
                        // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
                        TargetProperty.Name = "Image";
                        TargetProperty.Value = SourceProperty.Value;
                    }
                    break;

                case "ScrollBars":
                    // ScrollBars = System.Windows.Forms.ScrollBars.Both;
                    TargetProperty.Name = SourceProperty.Name;

                    if (Type == "RichTextBox")
                    {
                        TargetProperty.Value = "System.Windows.Forms.RichTextBoxScrollBars.";
                    }
                    else
                    {
                        TargetProperty.Value = "System.Windows.Forms.ScrollBars.";
                    }
                    switch (SourceProperty.Value)
                    {
                        default:
                        case "0":
                            TargetProperty.Value = TargetProperty.Value + "None";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "Horizontal";
                            break;

                        case "2":
                            TargetProperty.Value = TargetProperty.Value + "Vertical";
                            break;

                        case "3":
                            TargetProperty.Value = TargetProperty.Value + "Both";
                            break;
                    }
                    break;

                // SS tab
                case "TabOrientation":
                    TargetProperty.Name = "Alignment";
                    TargetProperty.Value = "System.Windows.Forms.TabAlignment.";
                    switch (SourceProperty.Value)
                    {
                        default:
                        case "0":
                            TargetProperty.Value = TargetProperty.Value + "Top";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "Bottom";
                            break;

                        case "2":
                            TargetProperty.Value = TargetProperty.Value + "Left";
                            break;

                        case "3":
                            TargetProperty.Value = TargetProperty.Value + "Right";
                            break;
                    }
                    break;

                // begin Listview

                // unsupported properties
                case "_ExtentX":
                case "_ExtentY":
                case "_Version":
                case "OLEDropMode":
                    ValidProperty = false;
                    break;

                // this.listView.View = System.Windows.Forms.View.List;
                case "View":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = "System.Windows.Forms.View.";
                    switch (SourceProperty.Value)
                    {
                        case "0":
                            TargetProperty.Value = TargetProperty.Value + "Details";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "LargeIcon";
                            break;

                        case "2":
                            TargetProperty.Value = TargetProperty.Value + "SmallIcon";
                            break;

                        case "3":
                        default:
                            TargetProperty.Value = TargetProperty.Value + "List";
                            break;
                    }
                    break;

                case "LabelEdit":
                case "LabelWrap":
                case "MultiSelect":
                case "HideSelection":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = GetBool(SourceProperty.Value);
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
                    ValidProperty = false;
                    break;

                // supported properties

                case "ControlBox":
                case "KeyPreview":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "ClientHeight":
                    TargetProperty.Name = "ClientSize";
                    TargetProperty.Value = "new System.Drawing.Size(" + GetSize("ClientHeight", "ClientWidth", SourcePropertyList) + ")";
                    break;

                case "ClientWidth":
                    // nothing, already processed by Height, Left
                    ValidProperty = false;
                    break;

                case "ClientLeft":
                case "ClientTop":
                    ValidProperty = false;
                    break;

                case "MaxButton":
                    TargetProperty.Name = "MaximizeBox";
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "MinButton":
                    TargetProperty.Name = "MinimizeBox";
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "WhatsThisHelp":
                    TargetProperty.Name = "HelpButton";
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "ShowInTaskbar":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                case "WindowList":
                    TargetProperty.Name = "MdiList";
                    TargetProperty.Value = GetBool(SourceProperty.Value);
                    break;

                // this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                // 0 - normal
                // 1 - minimized
                // 2 - maximized
                case "WindowState":
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = "System.Windows.Forms.FormWindowState.";
                    switch (SourceProperty.Value)
                    {
                        case "0":
                        default:
                            TargetProperty.Value = TargetProperty.Value + "Normal";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "Minimized";
                            break;

                        case "2":
                            TargetProperty.Value = TargetProperty.Value + "Maximized";
                            break;
                    }
                    break;

                case "StartUpPosition":
                    // 0 - manual
                    // 1 - center owner
                    // 2 - center screen
                    // 3 - windows default
                    TargetProperty.Name = "StartPosition";
                    TargetProperty.Value = "System.Windows.Forms.FormStartPosition.";
                    switch (SourceProperty.Value)
                    {
                        case "0":
                            TargetProperty.Value = TargetProperty.Value + "Manual";
                            break;

                        case "1":
                            TargetProperty.Value = TargetProperty.Value + "CenterParent";
                            break;

                        case "2":
                            TargetProperty.Value = TargetProperty.Value + "CenterScreen";
                            break;

                        case "3":
                        default:
                            TargetProperty.Value = TargetProperty.Value + "WindowsDefaultLocation";
                            break;
                    }
                    break;

                default:
                    TargetProperty.Name = SourceProperty.Name;
                    TargetProperty.Value = SourceProperty.Value;
                    TargetProperty.Valid = false;
                    break;
            }
            return ValidProperty;
        }

        public static bool ParseVariable(Variable SourceVariable, Variable TargetVariable)
        {
            TargetVariable.Scope = SourceVariable.Scope;
            TargetVariable.Name = SourceVariable.Name;
            TargetVariable.Type = VariableTypeConvert(SourceVariable.Type);

            return true;
        }

        public static bool ParseVariables(ArrayList SourceVariableList, ArrayList TargetVariableList)
        {
            // each property
            foreach (Variable SourceVariable in SourceVariableList)
            {
                var TargetVariable = new Variable();
                if (ParseVariable(SourceVariable, TargetVariable))
                {
                    TargetVariableList.Add(TargetVariable);
                }
            }
            return true;
        }

        private static void ControlListLoad()
        {
            _mControlList = new Hashtable();
            var Doc = new XmlDocument();
            XmlNode node;
            ControlListItem oItem;

            // get current directory
            string[] CommandLineArgs;
            CommandLineArgs = Environment.GetCommandLineArgs();
            // index 0 contain path and name of exe file
            var BinPath = Path.GetDirectoryName(CommandLineArgs[0].ToLower());
            var FileName = BinPath + @"\vb2c.xml";

            Doc.Load(FileName);
            // Select the node given
            node = Doc.DocumentElement.SelectSingleNode("/configuration/ControlList");
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

        private static void ConvertFont(ControlProperty SourceProperty, ControlProperty TargetProperty)
        {
            var FontName = string.Empty;
            var FontSize = 0;
            var FontCharSet = 0;
            var FontBold = false;
            var FontUnderline = false;
            var FontItalic = false;
            var FontStrikethrough = false;
            var Temp = string.Empty;
            //      BeginProperty Font
            //         Name            =   "Arial"
            //         Size            =   8.25
            //         Charset         =   238
            //         Weight          =   400
            //         Underline       =   0   'False
            //         Italic          =   0   'False
            //         Strikethrough   =   0   'False
            //      EndProperty

            foreach (ControlProperty oProperty in SourceProperty.PropertyList)
            {
                switch (oProperty.Name)
                {
                    case "Name":
                        FontName = oProperty.Value;
                        break;

                    case "Size":
                        FontSize = GetFontSizeInt(oProperty.Value);
                        break;

                    case "Weight":
                        //        If tLogFont.lfWeight >= FW_BOLD Then
                        //          bFontBold = True
                        //        Else
                        //          bFontBold = False
                        //        End If
                        // FW_BOLD = 700
                        FontBold = (int.Parse(oProperty.Value) >= 700);
                        break;

                    case "Charset":
                        FontCharSet = int.Parse(oProperty.Value);
                        break;

                    case "Underline":
                        FontUnderline = (int.Parse(oProperty.Value) != 0);
                        break;

                    case "Italic":
                        FontItalic = (int.Parse(oProperty.Value) != 0);
                        break;

                    case "Strikethrough":
                        FontStrikethrough = (int.Parse(oProperty.Value) != 0);
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

            TargetProperty.Name = "Font";
            TargetProperty.Value = "new System.Drawing.Font(" + FontName + ",";
            TargetProperty.Value = TargetProperty.Value + FontSize.ToString() + "F,";

            Temp = string.Empty;
            if (FontBold)
            {
                Temp = "System.Drawing.FontStyle.Bold";
            }
            if (FontItalic)
            {
                if (Temp != string.Empty) { Temp = Temp + " | "; }
                Temp = Temp + "System.Drawing.FontStyle.Italic";
            }
            if (FontUnderline)
            {
                if (Temp != string.Empty) { Temp = Temp + " | "; }
                Temp = Temp + "System.Drawing.FontStyle.Underline";
            }
            if (FontStrikethrough)
            {
                if (Temp != string.Empty) { Temp = Temp + " | "; }
                Temp = Temp + "System.Drawing.FontStyle.Strikeout";
            }
            if (Temp == string.Empty)
            {
                TargetProperty.Value = TargetProperty.Value + " System.Drawing.FontStyle.Regular,";
            }
            else
            {
                TargetProperty.Value = TargetProperty.Value + " ( " + Temp + " ),";
            }
            TargetProperty.Value = TargetProperty.Value + " System.Drawing.GraphicsUnit.Point, ";
            TargetProperty.Value = TargetProperty.Value + "((System.Byte)(" + FontCharSet.ToString() + ")));";
        }

        private static string GetBool(string Value)
        {
            if (int.Parse(Value) == 0)
            {
                return "false";
            }
            else
            {
                return "true";
            }
        }

        private static string GetColor(string Value)
        {
            Color color;

            if (Value.Length < 3)
            {
                var ColorValue = "0x" + Value;
                color = ColorTranslator.FromWin32(Convert.ToInt32(ColorValue, 16));
            }
            else if (Value.StartsWith("&"))
            {
                Value = Value
                    .Replace("&H", "")
                    .Replace("&", "")
                    ;
                color = ColorTranslator.FromHtml("#" + Value);
            }
            else
            {
                color = Color.FromArgb(Convert.ToInt32(Value));

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
        private static string GetControlIndexName(string TabName)
        {
            //  this.SSTab1.(Tab(1).Control(4) = "Option1(0)";
            var Start = 0;
            var End = 0;

            Start = TabName.IndexOf("(", 0);
            if (Start > -1)
            {
                End = TabName.IndexOf(")", 0);
                return TabName.Substring(0, Start) + TabName.Substring(Start + 1, End - Start - 1);
            }
            else
            {
                return TabName;
            }
        }

        private static int GetFontSizeInt(string Value)
        {
            var Position = 0;

            Position = Value.IndexOf(",", 0);
            if (Position > -1)
            {
                return int.Parse(Value.Substring(0, Position));
            }

            Position = Value.IndexOf(".", 0);
            if (Position > 0)
            {
                return int.Parse(Value.Substring(0, Position));
            }
            return int.Parse(Value);
        }

        private static string GetLocation(ArrayList PropertyList)
        {
            var Left = 0;
            var Top = 0;

            // each property
            foreach (ControlProperty oProperty in PropertyList)
            {
                if (oProperty.Name == "Left")
                {
                    Left = int.Parse(oProperty.Value);
                    if (Left < 0)
                    {
                        Left = 75000 + Left;
                    }
                    Left = Left / 15;
                }
                if (oProperty.Name == "Top")
                {
                    Top = int.Parse(oProperty.Value) / 15;
                }
            }
            // 616, 520
            return Left.ToString() + ", " + Top.ToString();
        }

        private static string GetSize(string Height, string Width, ArrayList PropertyList)
        {
            var HeightValue = 0;
            var WidthValue = 0;

            // each property
            foreach (ControlProperty oProperty in PropertyList)
            {
                if (oProperty.Name == Height)
                {
                    HeightValue = int.Parse(oProperty.Value) / 15;
                }
                if (oProperty.Name == Width)
                {
                    WidthValue = int.Parse(oProperty.Value) / 15;
                }
            }
            // 0, 120
            return WidthValue.ToString() + ", " + HeightValue.ToString();
        }

        private static string VariableTypeConvert(string SourceType)
        {
            string TargetType;

            switch (SourceType)
            {
                case "Long":
                    TargetType = "int";
                    break;

                case "Integer":
                    TargetType = "short";
                    break;

                case "Byte":
                    TargetType = "byte";
                    break;

                case "String":
                    TargetType = "string";
                    break;

                case "Boolean":
                    TargetType = "bool";
                    break;

                case "Currency":
                    TargetType = "decimal";
                    break;

                case "Single":
                    TargetType = "float";
                    break;

                case "Double":
                    TargetType = "double";
                    break;

                case "ADODB.Recordset":
                case "DAO.Recordset":
                case "Recordset":
                    TargetType = "DataReader";
                    break;

                default:
                    TargetType = SourceType;
                    break;
            }
            return TargetType;
        }
    }
}