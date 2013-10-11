using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
	public class OptionExplicitHandler : AbstractBlockHandler
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

            // Determine whether we've got an "OPTION EXPLICIT" statement
            if (!base.checkAtomTokenPattern(tokens, new string[] { "OPTION", "EXPLICIT" }, false))
                return null;

			// Pull content from token stream
			var numberOfTokensToRemove = 2;
			if (tokens.Count > numberOfTokensToRemove)
			{
				// Unless we've hit token stream end, next should be end-of-statement
				if (!base.isEndOfStatement(tokens, numberOfTokensToRemove))
					throw new Exception("No end-of-statement after \"OPTION EXPLICIT\" statement");
				numberOfTokensToRemove++;
			}
			tokens.RemoveRange(0, numberOfTokensToRemove);
			return new OptionExplicit();
        }
    }
}