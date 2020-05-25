using System.Collections;
using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    /// Summary description for Module.
    /// </summary>
    public class Module
    {
        public List<Control> ControlList { get; set; }
        public List<Enum> EnumList { get; set; }
        public string FileName { get; set; }
        public List<ControlProperty> FormPropertyList { get; }
        public List<string> ImageList { get; set; }
        public bool ImagesUsed { get; set; } = false;
        public bool MenuUsed { get; set; } = false;
        public string Name { get; set; }

        public List<Procedure> ProcedureList { get; set; }
        public List<Property> PropertyList { get; set; }
        public string Type { get; set; }

        public List<Variable> VariableList { get; set; }
        public string Version { get; set; }

        public string Comment { get; set; }

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