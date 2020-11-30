using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms.VisualStyles;
using MetX.VB6ToCSharp.CSharp;
using MetX.VB6ToCSharp.Interface;
using MetX.VB6ToCSharp.Structure;

namespace MetX.VB6ToCSharp.VB6
{
    /// <summary>
    ///     Summary description for Module.
    /// </summary>
    public class Module : AbstractBlock
    {
        public string Comment;
        public List<Control> ControlList;
        public List<Enum> EnumList;
        public string FileName;
        public List<ControlProperty> FormPropertyList;
        public List<string> ImageList;
        public bool ImagesUsed = false;
        public bool MenuUsed = false;
        public string Name;

        public List<Procedure> ProcedureList;
        public List<IAmAProperty> PropertyList;
        public string Type;

        public VariableList VariableList;
        public string Version;

        public Module(ICodeLine parent) : base(parent)
        {
            FormPropertyList = new List<ControlProperty>();
            ControlList = new List<Control>();
            ImageList = new List<string>();
            VariableList = new VariableList();
            PropertyList = new List<IAmAProperty>();
            ProcedureList = new List<Procedure>();
            EnumList = new List<Enum>();
            ResetIndentsRecursively = ModuleRIC;
        }

        public int ModuleRIC(int indentLevel)
        {
            foreach (var item in FormPropertyList)
            {
                item.ResetIndent(indentLevel + 1);
            }

            foreach (var item in ControlList)
            {
                item.ResetIndent(indentLevel + 1);
            }

            foreach (var item in ImageList) { }

            foreach (var item in VariableList)
            {
                item.ResetIndent(indentLevel + 1);
            }


            foreach (var item in PropertyList)
            {
                item.ResetIndent(indentLevel + 1);
            }
            
            foreach (var item in ProcedureList)
            {
                item.ResetIndent(indentLevel + 1);
            }

            foreach (var item in EnumList) { }

            return indentLevel;
        }

        public void ExamineAndAdjust()
        {
            foreach (var item in FormPropertyList)
            {
                item.Type = item.Type.ExamineAndAdjustLine(CurrentlyInArea.Class);
            }

            foreach (var item in ControlList)
            {
                item.Type = item.Type.ExamineAndAdjustLine(CurrentlyInArea.Class);
            }

            foreach (var item in ImageList)
            {
                // Not needed?
            }

            foreach (var item in VariableList)
            {
                item.Type = item.Type.ExamineAndAdjustLine(CurrentlyInArea.Class);
            }


            foreach (var item in PropertyList)
            {
                item.Type = item.Type.ExamineAndAdjustLine(CurrentlyInArea.Class);
            }

            foreach (var item in ProcedureList)
            {
                for (var i = 0; i < item.LineList.Count; i++)
                {
                    item.LineList[i] = item.LineList[i].ExamineAndAdjustLine(CurrentlyInArea.Class);
                }
            }

            foreach (var item in EnumList)
            {
                // Not needed?
            }

        }
    }
}