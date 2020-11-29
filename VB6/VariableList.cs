using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetX.Library;

namespace MetX.VB6ToCSharp.VB6
{
    public class VariableList : List<Variable>
    {
        public string GenerateCode(int indentLevel)
        {
            var sb = new StringBuilder();
            sb.AppendLine();
            foreach (var variable in this)
            {
                variable.ResetIndent(indentLevel);
                sb.AppendLine(variable.GenerateCode());
            }

            return sb.ToString();
        }

        public Variable this[string name]
        {
            get
            {
                var commonName = name.Trim().ToLower();
                foreach (var variable in this
                    .Where(variable => variable
                        .Name.Trim().ToLower() == commonName))
                {
                    return variable;
                }
                return new Variable
                {
                    Name = name,
                    Type = "object",
                };
            }
        }

        public bool Contains(string name)
        {
            if (name.IsEmpty())
                return false;

            var nameLower = name.Trim().ToLower();
            var result = this.Any(x => x.Name?.Trim().ToLower() == nameLower);
            return result;
        }
    }
}