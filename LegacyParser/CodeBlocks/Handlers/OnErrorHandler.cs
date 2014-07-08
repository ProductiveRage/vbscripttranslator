using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class OnErrorHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// Note: VBScript only supports "ON ERROR RESUME NEXT" and "ON ERROR GOTO 0" - it does not support specifying a label for the GOTO form
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            // Look for "ON ERROR.." form in tokens
            // - Define token matches with corresponding ICodeBlock type
            var matchPatterns = new Dictionary<string[], Func<int, ICodeBlock>>();
            matchPatterns.Add(
                new string[] { "ON", "ERROR", "RESUME", "NEXT" },
                lineIndex => new OnErrorResumeNext(lineIndex)
            );
            matchPatterns.Add(
                new string[] { "ON", "ERROR", "GOTO", "0" },
                lineIndex => new OnErrorGoto0(lineIndex)
            );
            // - Check for match
            int? tokensToRemove = null;
            ICodeBlock errorBlock = null;
            foreach (string[] matchPattern in matchPatterns.Keys)
            {
                if (base.checkAtomTokenPattern(tokens, matchPattern, false))
                {
                    errorBlock = matchPatterns[matchPattern](tokens[0].LineIndex);
                    tokensToRemove = matchPattern.Length;
                    break;
                }
            }
            if (tokensToRemove == null)
                return null;

            // Pull content from token stream
            if (tokens.Count > tokensToRemove)
            {
                // Unless we've hit token stream end, next should be end-of-statement
                if (!base.isEndOfStatement(tokens, tokensToRemove.Value))
                    throw new Exception("No end-of-statement after \"ON ERROR..\" statement");
                tokensToRemove++;
            }
            tokens.RemoveRange(0, tokensToRemove.Value);
            return errorBlock;
       }
    }
}
