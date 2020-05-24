using System.Collections;

namespace VB2C
{
    /// <summary>
    /// Summary description for Module.
    /// </summary>
    public class Module
    {
        public ArrayList ControlList { get; set; }
        public ArrayList EnumList { get; set; }
        public string FileName { get; set; }
        public ArrayList FormPropertyList { get; }
        public ArrayList ImageList { get; set; }
        public bool ImagesUsed { get; set; } = false;
        public bool MenuUsed { get; set; } = false;
        public string Name { get; set; }

        public ArrayList ProcedureList { get; set; }
        public ArrayList PropertyList { get; set; }
        public string Type { get; set; }

        public ArrayList VariableList { get; set; }
        public string Version { get; set; }

        public string Comment { get; set; }

        public Module()
        {
            FormPropertyList = new ArrayList();
            ControlList = new ArrayList();
            ImageList = new ArrayList();
            VariableList = new ArrayList();
            PropertyList = new ArrayList();
            ProcedureList = new ArrayList();
            EnumList = new ArrayList();
        }

        public void ControlAdd(Control oControl)
        {
            ControlList.Add(oControl);
        }

        public void FormPropertyAdd(ControlProperty oProperty)
        {
            FormPropertyList.Add(oProperty);
        }

        public void ProcedureAdd(Procedure oProcedure)
        {
            ProcedureList.Add(oProcedure);
        }

        public void PropertyAdd(Property oProperty)
        {
            PropertyList.Add(oProperty);
        }

        public void VariableAdd(Variable oVariable)
        {
            VariableList.Add(oVariable);
        }
    }
}