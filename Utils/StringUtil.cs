using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace me.fengyj.CommonLib.Utils {
    public class StringUtil {

        public static string TryAddOrGetNewNameIfDuplicated(ICollection<string> list, string name, int maxLength = int.MaxValue, string concatStr = "") {

            if (string.IsNullOrEmpty(name)) throw new ArgumentException("The parameter cannot be null nor empty.", nameof(name));
            
            var strTrimmed = name.Length < maxLength ? name : name[..maxLength];

            if (!list.Contains(strTrimmed)) {
                list.Add(strTrimmed);
                return strTrimmed;
            }
            else {
                var idx = 1;
                var newStr = strTrimmed;
                while (list.Contains(newStr)) {

                    var seqStr = $"{concatStr ?? string.Empty}{idx}";
                    var seqStrLen = seqStr.Length;
                    if (strTrimmed.Length + seqStrLen > maxLength)
                        newStr = strTrimmed[..(maxLength - seqStrLen)];
                    newStr += seqStr;
                    idx++;
                }
                list.Add(newStr);
                return newStr;
            }
        }

        /// <summary>
        /// Escape the characters: &lt;, &gt;, &quote; &apos; and &amp;.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string? EscapeForXML(string? str) {
            return System.Security.SecurityElement.Escape(str);
        }

        public static string? UnescapeForXML(string? str) {
            return str == null ? null : System.Security.SecurityElement.FromString($"<x>{str}</>")?.Text;
        }
    }
}
