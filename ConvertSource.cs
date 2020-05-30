using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    public static class ConvertSource
    {
        public static void ControlPropertyFont(ControlProperty sourceProperty, ControlProperty targetProperty)
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

            foreach (var property in sourceProperty.PropertyList)
            {
                switch (property.Name)
                {
                    case "Name":
                        fontName = property.Value;
                        break;

                    case "Size":
                        fontSize = Tools.GetFontSizeInt(property.Value);
                        break;

                    case "Weight":
                        //        If tLogFont.lfWeight >= FW_BOLD Then
                        //          bFontBold = True
                        //        Else
                        //          bFontBold = False
                        //        End If
                        // FW_BOLD = 700
                        fontBold = (int.Parse(property.Value) >= 700);
                        break;

                    case "Charset":
                        fontCharSet = int.Parse(property.Value);
                        break;

                    case "Underline":
                        fontUnderline = (int.Parse(property.Value) != 0);
                        break;

                    case "Italic":
                        fontItalic = (int.Parse(property.Value) != 0);
                        break;

                    case "Strikethrough":
                        fontStrikethrough = (int.Parse(property.Value) != 0);
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
                if (temp != string.Empty)
                {
                    temp += " | ";
                }

                temp += "System.Drawing.FontStyle.Italic";
            }

            if (fontUnderline)
            {
                if (temp != string.Empty)
                {
                    temp += " | ";
                }

                temp += "System.Drawing.FontStyle.Underline";
            }

            if (fontStrikethrough)
            {
                if (temp != string.Empty)
                {
                    temp += " | ";
                }

                temp += "System.Drawing.FontStyle.Strikeout";
            }

            if (temp == string.Empty)
            {
                targetProperty.Value += " System.Drawing.FontStyle.Regular,";
            }
            else
            {
                targetProperty.Value = targetProperty.Value + " ( " + temp + " ),";
            }

            targetProperty.Value += " System.Drawing.GraphicsUnit.Point, ";
            targetProperty.Value = targetProperty.Value + "((System.Byte)(" + fontCharSet.ToString() + ")));";
        }

        public static void ConvertLineOfCode(
            string originalLine,
            out string translatedLine,
            out string placeAtBottom,
            IAmAProperty sourceProperty)
        {
            var localSourceProperty = (Property) sourceProperty;
            placeAtBottom = string.Empty;
            var line = originalLine.Trim();
            translatedLine = line;

            if (translatedLine.Length > 0)
            {
                if (translatedLine.Trim().StartsWith("'"))
                {
                    translatedLine = "// " + translatedLine.Substring(1);
                    return;
                }

                if (localSourceProperty != null &&
                    (localSourceProperty.Direction == "Set" || localSourceProperty.Direction == "Let"))
                {
                    foreach (var parameter in localSourceProperty.Parameters)
                    {
                        translatedLine = translatedLine.Replace(parameter.Name, "value");
                    }
                }

                // vbNullString = string.empty
                if (translatedLine.Contains("vbNullString"))
                {
                    translatedLine = translatedLine.Replace("vbNullString", "string.Empty");
                }

                // Nothing = null
                if (translatedLine.Contains("Nothing"))
                {
                    translatedLine = line
                        .Replace("Set ", "")
                        .Replace("Nothing", "null");
                    translatedLine += ";";
                }

                if (Tools.WithCurrentlyIn.IsNotEmpty())
                {
                    if (translatedLine.StartsWith("."))
                        translatedLine = Tools.WithCurrentlyIn + translatedLine;
                    if (translatedLine.Contains(" ."))
                    {
                        var words = translatedLine.Split(' ');
                        for (var i = 0; i < words.Length; i++)
                        {
                            if (words[i].StartsWith("."))
                            {
                                words[i] = Tools.WithCurrentlyIn + words[i];
                            }
                        }

                        translatedLine = string.Join(" ", words);
                    }
                }

                // Begin With
                if (translatedLine.StartsWith("With "))
                {
                    Tools.WithCurrentlyIn = translatedLine.TokenAt(2).Trim();
                    translatedLine = "";
                }

                // Set
                if (translatedLine.Contains("Set "))
                {
                    translatedLine = translatedLine.Replace("Set ", " ");
                }

                if (translatedLine.Contains("Dim") && translatedLine.Contains("As"))
                {
                    // Dim x As string
                    // string x;
                    translatedLine = $"{translatedLine.TokenAt(4)} {translatedLine.TokenAt(2)} ";
                }

                // remark
                if (line[0] == '\'') // '
                {
                    translatedLine = translatedLine.Replace("'", "//");
                }

                // & to +
                if (translatedLine.Contains("&"))
                {
                    translatedLine = translatedLine.Replace("&", "+");
                }

                // Select Case
                if (translatedLine.Contains("Select Case"))
                {
                    translatedLine = translatedLine.Replace("Select Case", "switch");
                }

                // End Select
                if (translatedLine.Contains("End Select"))
                {
                    translatedLine = translatedLine.Replace("End Select", "}");
                }

                // _
                if (translatedLine.Contains(" _"))
                {
                    translatedLine = translatedLine.Replace(" _", "\r\n");
                }

                // If
                if (translatedLine.Contains("If "))
                {
                    translatedLine = translatedLine.Replace("If ", "if ( ");
                }

                // Not
                if (translatedLine.Contains("Not "))
                {
                    translatedLine = translatedLine.Replace("Not ", "! ");
                }

                // then
                if (translatedLine.Contains(" Then"))
                {
                    translatedLine = translatedLine.Replace(" Then", " )\r\n" + ConvertCode.Indent3 + "{\r\n");
                }

                // else
                if (translatedLine.Contains("Else"))
                {
                    translatedLine = translatedLine.Replace("Else",
                        "}\r\n" + ConvertCode.Indent3 + "else\r\n" + ConvertCode.Indent3 + "{");
                }

                // End if
                if (translatedLine.Contains("End If"))
                {
                    translatedLine = translatedLine.Replace("End If", "}");
                }

                // Unload Me
                if (translatedLine.Contains("Unload Me"))
                {
                    translatedLine = translatedLine.Replace("Unload Me", "Close()");
                }

                // .Caption
                if (translatedLine.Contains(".Caption"))
                {
                    translatedLine = translatedLine.Replace(".Caption", ".Text");
                }

                // True
                if (translatedLine.Contains("True"))
                {
                    translatedLine = translatedLine.Replace("True", "true");
                }

                // False
                if (translatedLine.Contains("False"))
                {
                    translatedLine = translatedLine.Replace("False", "false");
                }

                // New
                if (line.Contains("If ")
                    && line.Contains("Then")
                    && line.TokensAfter(1, "Then").Trim().Length > 0)
                {
                    translatedLine = translatedLine.Replace("New", "new");
                }

                // New
                if (translatedLine.Contains("New "))
                {
                    translatedLine = translatedLine.Replace("New ", "new ");
                    if (!translatedLine.Contains("("))
                        translatedLine += "();";
                }

                if (translatedLine.Contains("On Error Resume Next"))
                {
                    placeAtBottom = @"
        }
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
";
                    translatedLine = "try\r\n{\r\n";
                }
            }

            if (translatedLine.IsNotEmpty())
            {
                translatedLine = Massage.Now(translatedLine);
            }
        }

        public static bool Module(Module sourceModule, Module targetModule)
        {
            // module name
            targetModule.Name = sourceModule.Name;
            // file name
            targetModule.FileName = Path.GetFileNameWithoutExtension(sourceModule.FileName) + ".cs";
            // type
            targetModule.Type = sourceModule.Type;
            // version
            targetModule.Version = sourceModule.Version;
            // process own properties - forms
            ConvertSourceModuleProperties(targetModule, sourceModule.FormPropertyList, targetModule.FormPropertyList);
            // process controls - form
            Controls(targetModule, sourceModule.ControlList, targetModule.ControlList);

            // special exception for menu
            if (targetModule.MenuUsed)
            {
                // add main menu control
                var control = new Control
                {
                    Name = "MainMenu",
                    Owner = targetModule.Name,
                    Type = "MainMenu",
                    Valid = true,
                    InvisibleAtRuntime = true
                };
                targetModule.ControlList.Insert(0, control);
                foreach (var oMenuControl in targetModule.ControlList)
                    if ((oMenuControl.Type == "MenuItem") && (oMenuControl.Owner == targetModule.Name))
                        // rewrite previous owner
                        oMenuControl.Owner = control.Name;
            }

            var tempControlList = new List<Control>();
            var tabControlIndex = 0;

            // check for TabDlg.SSTab
            foreach (var targetControl in targetModule.ControlList)
            {
                Control tabPage = null;
                var index = 0;
                if ((targetControl.Type == "TabControl") && (targetControl.Valid))
                {
                    // each property
                    foreach (var targetProperty in targetControl.PropertyList)
                    {
                        Console.WriteLine(targetProperty.Name);

                        if (targetProperty.Name.Contains($"TabCaption({index})"))
                        {
                            // new tab
                            tempControlList.Add(new Control
                            {
                                Type = "TabPage",
                                Name = "tabPage" + index.ToString(),
                                Owner = targetControl.Name,
                                Container = true,
                                Valid = true,
                                InvisibleAtRuntime = false,
                                // add some necessary properties
                                PropertyList = new List<ControlProperty>
                                {
                                    new ControlProperty(2)
                                    {
                                        Name = "Location", Value = "new System.Drawing.Point(4, 22)", Valid = true
                                    },
                                    new ControlProperty(2)
                                    {
                                        Name = "Size", Value = "new System.Drawing.Size(477, 374)", Valid = true
                                    },
                                    new ControlProperty(2)
                                    {
                                        Name = "Text", Value = targetProperty.Value, Valid = true
                                    },
                                    new ControlProperty(2)
                                    {
                                        Name = "TabIndex", Value = index.ToString(), Valid = true
                                    }
                                }
                            });
                            index++;
                        }

                        // Control = change owner of control to current tab
                        //      this.SSTab1.(Tab(0).Control(0) = "ImageControl";
                        if (targetProperty.Name.Contains(".Control("))
                        {
                            if (!targetProperty.Name.Contains("Enable"))
                            {
                                var tabName = targetProperty.Value.Substring(1, targetProperty.Value.Length - 2);
                                tabName = Tools.GetControlIndexName(tabName);
                                // search for "targetProperty.Value" control
                                // and replace owner of this control to current tab
                                foreach (var oNewOwner in targetModule.ControlList)
                                {
                                    if ((oNewOwner.Name == tabName) && (!oNewOwner.InvisibleAtRuntime))
                                    {
                                        oNewOwner.Owner = tabPage.Name;
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
                foreach (var control in tempControlList)
                {
                    targetModule.ControlList.Insert(tabControlIndex + position, control);
                    position++;
                }
            }

            // process enums
            Enums(sourceModule.EnumList, targetModule.EnumList);

            // process variables
            Variables(sourceModule.VariableList, targetModule.VariableList);

            // process properties
            ClassProperties(sourceModule.PropertyList, targetModule.PropertyList);

            // process procedures
            Procedures(sourceModule, targetModule.ProcedureList);

            return true;
        }

        public static bool ConvertSourceModuleProperties(
            Module targetModule,
            List<ControlProperty> sourcePropertyList,
            List<ControlProperty> targetPropertyList)
        {
            foreach (var sourceProperty in sourcePropertyList)
            {
                var targetProperty = new ControlProperty(2);
                if (!ControlProperties(targetModule.Type, sourceProperty, targetProperty, sourcePropertyList)) 
                    continue;

                if (targetProperty.Name == "BackgroundImage" || targetProperty.Name == "Icon")
                    targetModule.ImagesUsed = true;

                targetPropertyList.Add(targetProperty);
            }
            return true;
        }

        public static bool ClassProperties(
            List<IAmAProperty> sourceProperties, List<IAmAProperty> targetProperties)
        {
            foreach (var sourceProperty in sourceProperties)
            {
                var localSourceProperty = (Property) sourceProperty;
                CSharpProperty targetProperty = null;

                targetProperty = targetProperties.Cast<CSharpProperty>()
                    .FirstOrDefault(x => string
                        .Equals(x.Name, sourceProperty.Name, StringComparison.CurrentCultureIgnoreCase));

                if (targetProperty == null)
                {
                    targetProperty = new CSharpProperty(2)
                    {
                        Name = sourceProperty.Name.Trim(),
                        Comment = sourceProperty.Comment,
                        Scope = sourceProperty.Scope,
                        Type = Tools.VariableTypeConvert(sourceProperty.Type),
                    };
                    targetProperties.Add(targetProperty);
                }

                targetProperty.ConvertSourcePropertyParts(sourceProperty);
            }

            return true;
        }

        public static bool ControlProperties(
            Module module,
            Control control,
            List<ControlProperty> sourcePropertyList,
            List<ControlProperty> targetPropertyList)
        {
            // each property
            foreach (var sourceProperty in sourcePropertyList)
            {
                if (sourceProperty.Name == "Index")
                {
                    // Index           =   3
                    control.Name += sourceProperty.Value;
                }
                else
                {
                    var targetProperty = new ControlProperty(2);
                    if (ControlProperties(control.Type, sourceProperty, targetProperty, sourcePropertyList))
                    {
                        if (targetProperty.Name == "Image")
                        {
                            module.ImagesUsed = true;
                        }

                        targetPropertyList.Add(targetProperty);
                    }
                }
            }

            return true;
        }

        public static bool Controls(
            Module module,
            List<Control> sourceControlList,
            List<Control> targetControlList)
        {
            var type = string.Empty;

            foreach (var sourceControl in sourceControlList)
            {
                var targetControl = new Control
                {
                    Name = sourceControl.Name,
                    Owner = sourceControl.Owner,
                    Container = sourceControl.Container,
                    Valid = true
                };

                type = sourceControl.Type;

                targetControl.Type = type;
                ControlProperties(module, targetControl, sourceControl.PropertyList, targetControl.PropertyList);

                targetControlList.Add(targetControl);
            }

            return true;
        }

        public static bool Enums(List<Enum> sourceEnums, List<Enum> targetEnums)
        {
            foreach (var sourceEnum in sourceEnums)
                targetEnums.Add(sourceEnum);

            return true;
        }

        public static bool Procedures(Module sourceModule, List<Procedure> targetProcedures)
        {
            const string indent6 = "      ";

            foreach (var sourceProcedure in sourceModule.ProcedureList)
            {
                var targetProcedure = new Procedure
                {
                    Name = sourceProcedure.Name,
                    Scope = sourceProcedure.Scope,
                    Comment = sourceProcedure.Comment,
                    Type = sourceProcedure.Type,
                    ReturnType = Tools.VariableTypeConvert(sourceProcedure.ReturnType),
                    ParameterList = sourceProcedure.ParameterList
                };

                // lines
                foreach (var originalLine in sourceProcedure.LineList)
                {
                    ConvertLineOfCode(originalLine, out var convertedLine, out var placeAtBottom, null);

                    targetProcedure.LineList.Add(convertedLine);
                    if (placeAtBottom.IsNotEmpty())
                        targetProcedure.BottomLineList.Add(placeAtBottom);
                }

                Massage.DetermineWhichLinesGetASemicolon(targetProcedure.LineList);
                Massage.DetermineWhichLinesGetASemicolon(targetProcedure.BottomLineList);

                targetProcedures.Add(targetProcedure);
            }

            return true;
        }

        public static bool ControlProperties(string type,
            ControlProperty sourceProperty,
            ControlProperty targetProperty,
            List<ControlProperty> sourcePropertyList)
        {
            var validProperty = true;
            targetProperty.Valid = true;

            switch (sourceProperty.Name)
            {
                // not used
                case "Appearance":
                case "ScaleHeight":
                case "ScaleWidth":
                case "Style": // button
                case "BackStyle": //label
                case "IMEMode":
                case "WhatsThisHelpID":
                case "Mask": // maskedit
                case "PromptChar": // maskedit
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
                            targetProperty.Value += "TopLeft";
                            break;

                        case "1":
                            targetProperty.Value += "TopRight";
                            break;

                        case "2":
                        default:
                            targetProperty.Value += "TopCenter";
                            break;
                    }

                    break;

                case "BackColor":
                case "ForeColor":
                    if (type != "ImageList")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = Tools.GetColor(sourceProperty.Value);
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
                                targetProperty.Value += "None";
                                break;

                            default:
                            case "1":
                                targetProperty.Value += "FixedSingle";
                                break;

                            case "2":
                                targetProperty.Value += "Sizable";
                                break;

                            case "3":
                                targetProperty.Value += "FixedDialog";
                                break;

                            case "4":
                                targetProperty.Value += "FixedToolWindow";
                                break;

                            case "5":
                                targetProperty.Value += "SizableToolWindow";
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
                                targetProperty.Value += "None";
                                break;

                            case "1":
                                targetProperty.Value += "FixedSingle";
                                break;

                            case "2":
                            default:
                                targetProperty.Value += "Fixed3D";
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
                    targetProperty.Value = "new System.Drawing.Size(" + Tools.GetSize("Height", "Width", sourcePropertyList) +
                                           ")";
                    break;

                // this.cmdExit.Location = new System.Drawing.Point(616, 520);
                case "Left":
                    if ((type != "ImageList") && (type != "Timer"))
                    {
                        targetProperty.Name = "Location";
                        targetProperty.Value = "new System.Drawing.Point(" + Tools.GetLocation(sourcePropertyList) + ")";
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
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "WordWrap":
                    if (type == "Text")
                    {
                        targetProperty.Name = sourceProperty.Name;
                        targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    }
                    else
                    {
                        validProperty = false;
                    }

                    break;

                case "Font":
                    ControlPropertyFont(sourceProperty, targetProperty);
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
                            targetProperty.Value = Tools.GetBool(sourceProperty.Value);
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
                                    targetProperty.Value += "Unchecked";
                                    break;

                                case "1":
                                    targetProperty.Value += "Checked";
                                    break;

                                case "2":
                                    targetProperty.Value += "Indeterminate";
                                    break;
                            }

                            break;

                        default:
                            targetProperty.Value += "Both";
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
                        targetProperty.Value = Tools.GetBool(sourceProperty.Value);
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
                            targetProperty.Value += "None";
                            break;

                        case "1":
                            targetProperty.Value += "Horizontal";
                            break;

                        case "2":
                            targetProperty.Value += "Vertical";
                            break;

                        case "3":
                            targetProperty.Value += "Both";
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
                            targetProperty.Value += "Top";
                            break;

                        case "1":
                            targetProperty.Value += "Bottom";
                            break;

                        case "2":
                            targetProperty.Value += "Left";
                            break;

                        case "3":
                            targetProperty.Value += "Right";
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
                            targetProperty.Value += "Details";
                            break;

                        case "1":
                            targetProperty.Value += "LargeIcon";
                            break;

                        case "2":
                            targetProperty.Value += "SmallIcon";
                            break;

                        case "3":
                        default:
                            targetProperty.Value += "List";
                            break;
                    }

                    break;

                case "LabelEdit":
                case "LabelWrap":
                case "MultiSelect":
                case "HideSelection":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
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
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "ClientHeight":
                    targetProperty.Name = "ClientSize";
                    targetProperty.Value = "new System.Drawing.Size(" + Tools.GetSize("ClientHeight", "ClientWidth", sourcePropertyList) + ")";
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
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "MinButton":
                    targetProperty.Name = "MinimizeBox";
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "WhatsThisHelp":
                    targetProperty.Name = "HelpButton";
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "ShowInTaskbar":
                    targetProperty.Name = sourceProperty.Name;
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
                    break;

                case "WindowList":
                    targetProperty.Name = "MdiList";
                    targetProperty.Value = Tools.GetBool(sourceProperty.Value);
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
                            targetProperty.Value += "Normal";
                            break;

                        case "1":
                            targetProperty.Value += "Minimized";
                            break;

                        case "2":
                            targetProperty.Value += "Maximized";
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
                            targetProperty.Value += "Manual";
                            break;

                        case "1":
                            targetProperty.Value += "CenterParent";
                            break;

                        case "2":
                            targetProperty.Value += "CenterScreen";
                            break;

                        case "3":
                        default:
                            targetProperty.Value += "WindowsDefaultLocation";
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

        public static bool Variable(Variable sourceVariable, Variable targetVariable)
        {
            targetVariable.Scope = sourceVariable.Scope;
            targetVariable.Name = sourceVariable.Name;
            targetVariable.Type = Tools.VariableTypeConvert(sourceVariable.Type);

            return true;
        }

        public static bool Variables(IList<Variable> sourceVariableList, IList<Variable> targetVariableList)
        {
            // each property
            foreach (var sourceVariable in sourceVariableList)
            {
                var targetVariable = new Variable();
                if (Variable(sourceVariable, targetVariable))
                {
                    targetVariableList.Add(targetVariable);
                }
            }

            return true;
        }
    }
}