using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class ExitHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            foreach (ExitStatement.ExitableStatementType exitType in Enum.GetValues(typeof(ExitStatement.ExitableStatementType)))
            {
                string[] matchPattern = new string[] { "EXIT", exitType.ToString() };
                if (base.checkAtomTokenPattern(tokens, matchPattern, false))
                {
                    var requireAnEndOfStatementToken = (tokens.Count > matchPattern.Length);
                    if (requireAnEndOfStatementToken)
                    {
                        if (!(tokens[matchPattern.Length] is AbstractEndOfStatementToken))
                            throw new Exception("EXIT statement wasn't followed by end-of-statement token");
                    }
                    tokens.RemoveRange(0, matchPattern.Length);
                    if (requireAnEndOfStatementToken)
                        tokens.RemoveRange(0, 1);
                    return new ExitStatement(exitType);
                }
            }
            return null;
        }
    }
}
