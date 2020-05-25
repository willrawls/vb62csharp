using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Module.
    /// </summary>
    public class Module
    {
        public List<Control> ControlList;
        public List<Enum> EnumList;
        public string FileName;
        public List<ControlProperty> FormPropertyList;
        public List<string> ImageList;
        public bool ImagesUsed = false;
        public bool MenuUsed = false;
        public string Name;

        public List<Procedure> ProcedureList;
        public List<Property> PropertyList;
        public string Type;

        public List<Variable> VariableList;
        public string Version;

        public string Comment;

        public Module()
        {
            FormPropertyList = new List<ControlProperty>();
            ControlList = new List<Control>();
            ImageList = new List<string>();
            VariableList = new List<Variable>();
            PropertyList = new List<Property>();
            ProcedureList = new List<Procedure>();
            EnumList = new List<Enum>();
        }

        public void ControlAdd(Control control)
        {
            ControlList.Add(control);
        }

        public void FormPropertyAdd(ControlProperty property)
        {
            FormPropertyList.Add(property);
        }

        public void ProcedureAdd(Procedure procedure)
        {
            ProcedureList.Add(procedure);
        }

        public void PropertyAdd(Property property)
        {
            PropertyList.Add(property);
        }

        public void VariableAdd(Variable variable)
        {
            VariableList.Add(variable);
        }
    }
}