using System.Collections.Generic;

namespace MetX.VB6ToCSharp
{
    /// <summary>
    ///     Summary description for Module.
    /// </summary>
    public class Module
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

        public List<Variable> VariableList;
        public string Version;

        public Module()
        {
            FormPropertyList = new List<ControlProperty>();
            ControlList = new List<Control>();
            ImageList = new List<string>();
            VariableList = new List<Variable>();
            PropertyList = new List<IAmAProperty>();
            ProcedureList = new List<Procedure>();
            EnumList = new List<Enum>();
        }
    }
}