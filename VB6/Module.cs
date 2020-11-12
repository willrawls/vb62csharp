using System.Collections.Generic;
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
        }
    }
}