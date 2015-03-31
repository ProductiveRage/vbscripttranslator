using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    /// <summary>
    /// TODO: This interpretation of CONST is a hack, it treats "CONST a = 1" as "DIM a: a = 1" which means that the CONST nature of "a" is not
    /// enforced in the translated code (so "a" may be altered without error, whereas in the VBScript source such an operation would result in
    /// a runtime error). I've just included this for now since I think it's the last structure that is not supported for parsing (having run
    /// it on thousands of lines of real VBScript so far). I need to revisit this and, probably, represent CONST values in ScopeAccessInformation
    /// data that the translation process uses, so that it can be ensure that the defined value is not altered (or, at least, that a runtime
    /// error is raised - which, as a runtime error, must be trappable by ON ERROR RESUME NEXT).
    /// </summary>
    public class ConstHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            if (!base.checkAtomTokenPattern(tokens, new string[] { "CONST" }, false))
                return null;

            // The only acceptable values for a CONST value are number/string literal and a subset of the builtin VBScript values (such as
            // true, false and empty but not including constants such as vbObjectError)
            var acceptableBuiltinValues = new[] { "true", "false", "empty", "null", "nothing" };

            NameToken name;
            if ((tokens.Count > 4)
            && (tokens[1] is NameToken)
            && (tokens[2] is OperatorToken)
            && (tokens[2].Content == "="))
            {
                if ((tokens[3] is NumericValueToken)
                || (tokens[3] is StringToken)
                || ((tokens[3] is BuiltInValueToken) && acceptableBuiltinValues.Contains(tokens[3].Content, StringComparer.OrdinalIgnoreCase)))
                    name = (NameToken)tokens[1];
                else
                    throw new ArgumentException("Invalid input - encountered invalid CONST statement (expected literal constant)");
            }
            else
                throw new ArgumentException("Invalid input - encountered invalid CONST statement");

            // Due to the hack that is being performed here for now (treating "CONST a = 1" as "DIM a: a = 1"), we only need to remove the first
            // token (the "CONST") and leave the remaining tokens to be identified as a ValueSettingStatement.
            tokens.RemoveAt(0);
            return new DimStatement(new[] {
                new DimStatement.DimVariable(name, dimensions: null)
            });
        }
    }
}
