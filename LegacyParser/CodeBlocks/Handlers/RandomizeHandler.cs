using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
	public class RandomizeHandler : AbstractBlockHandler
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

			// Determine whether we've got a "RANDOMIZE" statement
			if (!base.checkAtomTokenPattern(tokens, new string[] { "RANDOMIZE" }, false))
				return null;

			// Try to grab the tokens used to declare the seed (this is optional - if not specified, VBScript uses a time-based value). This
			// section will throw an exception if invalid tokens are encountered.
			var lineIndex = tokens[0].LineIndex; // We'll need this value, so get it before throwing any tokens away
			int tokensProcessed = 1;
			var seedTokens = new List<IToken>();
			for (int index = 1; index < tokens.Count; index++)
			{
				if (base.isEndOfStatement(tokens, index))
				{
					tokensProcessed++;
					break;
				}
				seedTokens.Add(base.getToken_AtomOrDateStringLiteralOnly(tokens, index));
				tokensProcessed++;
			}
			if ((seedTokens.Count == 2) && (seedTokens[0] is OpenBrace) && (seedTokens[1] is CloseBrace))
			{
				// Randomize may have optional brackets in its invocation - if it has a seed argument then this will be fine but if there
				// is no seed and there are brackets then the translator will get confused, so just strip them out in that case
				seedTokens.Clear();
			}

			// Pull processed tokens from stream and return statement
			tokens.RemoveRange(0, tokensProcessed);
			Expression seedIfAny = (seedTokens.Count == 0 ? null : new Expression(seedTokens));
			return new RandomizeStatement(lineIndex, seedIfAny: seedIfAny);
		}
	}
}