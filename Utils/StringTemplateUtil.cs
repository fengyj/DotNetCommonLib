using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Utils {
    /// <summary>
    /// replace the variable names in the template via Regex to get the format string, 
    /// and leverage the string.format function to return the text with the format string and variable values.
    /// </summary>
    public class StringTemplateUtil {

        /// <summary>
        /// the template with the default variable name standard. like $abc, _def, Xyz, i3, size.X, etc.
        /// </summary>
        /// <example>
        /// <![CDATA[
        /// template:  The $a's value is {$a,2}, __b's value is {__b:yyyyMMdd}, and c5's is {c5,-6:N2}.
        /// text:      The $a's value is  A, __b's value is 20240512, and c5's is 20.12 .
        /// variables: { a: "A", __b: 2024-05-12, c5: 20.1234 }
        /// ]]>
        /// </example>
        public static readonly StringTemplateUtil Default = new StringTemplateUtil("[\\$_A-Za-z][\\$_\\.\\w]*");

        private readonly Regex regex;

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="variableNameRegex">regex of the variable name</param>
        public StringTemplateUtil([StringSyntax("Regex")] string variableNameRegex) {
            this.regex = new Regex($"\\{{(?<id>{variableNameRegex})(,-?\\d+)?(:[^\\{{}}]+)?}}", RegexOptions.Compiled);
        }

        public string GetText(string template, Func<string, object?> valueGetter) {

            var vars = new List<string>();
            var evaluator = new MatchEvaluator(match => {
                var varName = match.Groups["id"].Value;
                vars.Add(varName);
                return match.Value.Replace(varName, (vars.Count - 1).ToString());
            });

            var fmt = this.regex.Replace(template, evaluator);
            return vars.Count == 0 ? template : string.Format(fmt, vars.Select(v => valueGetter(v)).ToArray());
        }

        public string GetText(string template, Dictionary<string, object?> vars, bool nullIfVariableNotExists = false) {
            if (nullIfVariableNotExists)
                return this.GetText(template, v => vars.TryGetValue(v, out var o) ? o : null);
            else
                return this.GetText(template, v => vars[v]);
        }
    }
}
