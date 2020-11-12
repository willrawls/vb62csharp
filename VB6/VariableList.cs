using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                // All public cause I really don't like anything but public stuff
                sb.AppendLine(variable.GenerateCode(indentLevel + 1));
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
            var nameLower = name.ToLower();
            var result = this.Any(x => x.Name.ToLower() == nameLower);
            return result;
        }
    }
}