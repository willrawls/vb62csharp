using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp
{
    // parse VB6 properties and values to C#
    public static class Tools
    {
        public static string WithCurrentlyIn = "";

        // When line starts with X
        //      Replace all instances of Y with Z
        //      Append A

        public static string GetBool(string value)
        {
            return value.IsEmpty() || (int.Parse(value) == 0)
                ? "false" 
                : "true";
        }

        public static string GetColor(string value)
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
        public static string GetControlIndexName(string tabName)
        {
            //  this.SSTab1.(Tab(1).Control(4) = "Option1(0)";
            var start = 0;
            var end = 0;

            start = tabName.IndexOf("(");
            if (start > -1)
            {
                end = tabName.IndexOf(")");
                return tabName.Substring(0, start) + tabName.Substring(start + 1, end - start - 1);
            }
            else
            {
                return tabName;
            }
        }

        public static int GetFontSizeInt(string value)
        {
            var position = 0;

            position = value.IndexOf(",");
            if (position > -1)
            {
                return int.Parse(value.Substring(0, position));
            }

            position = value.IndexOf(".");
            if (position > 0)
            {
                return int.Parse(value.Substring(0, position));
            }

            return int.Parse(value);
        }

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
            bytesToRead += (header[1] * 0x100);
            bytesToRead += (header[2] * 0x10000);
            bytesToRead += (header[3] * 0x1000000);
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

        public static string GetLocation(List<ControlProperty> propertyList)
        {
            var left = 0;
            var top = 0;

            // each property
            foreach (var property in propertyList)
            {
                if (property.Name == "Left")
                {
                    left = int.Parse(property.Value);
                    if (left < 0)
                    {
                        left = 75000 + left;
                    }

                    left /= 15;
                }

                if (property.Name == "Top")
                {
                    top = int.Parse(property.Value) / 15;
                }
            }

            // 616, 520
            return left.ToString() + ", " + top.ToString();
        }

        public static string GetSize(string height, string width, List<ControlProperty> propertyList)
        {
            var heightValue = 0;
            var widthValue = 0;

            // each property
            foreach (var property in propertyList)
            {
                if (property.Name == height)
                {
                    heightValue = int.Parse(property.Value) / 15;
                }

                if (property.Name == width)
                {
                    widthValue = int.Parse(property.Value) / 15;
                }
            }

            // 0, 120
            return widthValue.ToString() + ", " + heightValue.ToString();
        }

        public static string Indent(int level)
        {
            return level < 1
                ? string.Empty
                : new string(' ', level * 4);
        }

        public static string VariableTypeConvert(string sourceType)
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

        public static string Blockify(string blockName, int indentLevel, 
            string preBlock,  string postBlock, 
            Func<StringBuilder, string> action)
        {
            var indentation = Indent(indentLevel);
            var block = new StringBuilder();
            block.AppendLine(indentation + blockName.Trim());
            block.AppendLine(indentation + preBlock);
            var blockLines = action.Invoke(block);
            var lines = blockLines.Indent(indentLevel + 1);
            block.AppendLine(lines);
            block.AppendLine(indentation + postBlock);
            var code = block.ToString();
            return code;
        }
    }
}