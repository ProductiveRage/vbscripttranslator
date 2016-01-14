using System;
using System.Text;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
    public static class String_Extensions
    {
        /// <summary>
        /// Although VBScript only supports escaping of quotes, we need to run the full gamut when considering generating C# string literals since ANY
        /// content may be returned from a VBScriptNameRewriter or TempValueNameGenerator call
        /// </summary>
        public static string ToLiteral(this string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            var literal = new StringBuilder(input.Length + 2);
            literal.Append("\"");
            foreach (var c in input)
            {
                switch (c)
                {
                    case '\"': literal.Append("\\\""); break;
                    case '\\': literal.Append(@"\\"); break;
                    case '\0': literal.Append(@"\0"); break;
                    case '\a': literal.Append(@"\a"); break;
                    case '\b': literal.Append(@"\b"); break;
                    case '\f': literal.Append(@"\f"); break;
                    case '\n': literal.Append(@"\n"); break;
                    case '\r': literal.Append(@"\r"); break;
                    case '\t': literal.Append(@"\t"); break;
                    case '\v': literal.Append(@"\v"); break;
                    default:
                        if (char.IsControl(c))
                        {
                            literal.Append(@"\u");
                            literal.Append(((int)c).ToString("x4"));
                        }
                        else
                            literal.Append(c);
                        break;
                }
            }
            literal.Append("\"");
            return literal.ToString();
        }
    }
}
