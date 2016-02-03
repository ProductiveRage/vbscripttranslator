using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
	/// <summary>
	/// This will handle tokens only if they contain no keywords - if they do then they should have been picked up by one of the other handlers (unless
	/// these keywords are properties of an object, in which case it is valid - eg. Response.End is valid despite "End" being a VBScript keyword). For
	/// this reason, this should always be the last-resort handler.
	/// </summary>
	public class StatementHandler : AbstractBlockHandler
	{
		/// <summary>
		/// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
		/// </summary>
		public override ICodeBlock Process(List<IToken> tokens)
		{
			// Input validation
			if (tokens == null)
				throw new ArgumentNullException("tokens");
			if (tokens.Count == 0)
				return null;

			int indexRemoveTo = -1;
			List<IToken> tokensToExtract = new List<IToken>();
			for (int index = 0; index < tokens.Count; index++)
			{
				if (tokens[index] == null)
					throw new ArgumentException("Encountered null token in stream");
				if (tokens[index] is AtomToken)
				{
					IToken prevToken = (index == 0 ? null : tokens[index - 1]);
					if ((prevToken == null) || (!(prevToken is AtomToken)) || (prevToken.Content != "."))
					{
						// This is an AtomToken that does not appear to be a property of an object, so we need to ensure it's not a keyword that
						// should have been handled already by this point
						if (((AtomToken)tokens[index]).IsMustHandleKeyWord)
							throw new Exception("Encountered must-handle keyword in statement content, this should have been handled by a previous AbstractBlockHandler: \"" + tokens[index].Content + "\", line " + (tokens[index].LineIndex + 1) + " (this often indicates a mismatched block terminator, such as an END SUB when an END FUNCTION was expected)");
					}
				}
				if (tokens[index] is AbstractEndOfStatementToken)
				{
					indexRemoveTo = index;
					break;
				}
				tokensToExtract.Add(tokens[index]);
			}
			
			if (indexRemoveTo == -1)
				tokens.Clear();
			else
				tokens.RemoveRange(0, indexRemoveTo + 1);

			return ReturnAppropriateStatement(tokensToExtract);
		}

		private ICodeBlock ReturnAppropriateStatement(IEnumerable<IToken> tokens)
		{
			var hasCallPrefix = false;
			var isSetStatement = false;
			var initialTokens = tokens.ToList();
			var firstTokenAsKeyword = initialTokens.FirstOrDefault() as KeyWordToken;
			if (firstTokenAsKeyword != null)
			{
				bool cullFirstToken;
				if (firstTokenAsKeyword.Content.Equals("CALL", StringComparison.InvariantCultureIgnoreCase))
				{
					cullFirstToken = true;
					hasCallPrefix = true;
				}
				else if (firstTokenAsKeyword.Content.Equals("LET", StringComparison.InvariantCultureIgnoreCase))
					cullFirstToken = true;
				else if (firstTokenAsKeyword.Content.Equals("SET", StringComparison.InvariantCultureIgnoreCase))
				{
					cullFirstToken = true;
					isSetStatement = true;
				}
				else
					cullFirstToken = true;
				if (cullFirstToken)
					initialTokens = initialTokens.Skip(1).ToList();
			}

			var bracketCount = 0;
			var valueToSetTokens = new List<IToken>();
			var expressionTokens = new List<IToken>();
			var inExpressionContent = false;
			for (var index = 0; index < initialTokens.Count; index++)
			{
				var token = initialTokens[index];
				if ((token is ComparisonOperatorToken) & (token.Content == "=") && (bracketCount == 0) && !inExpressionContent)
				{
					// Taken an equals sign to indicate the break between a value-to-set and expression-to-set in a value-setting-
					// statement (eg. "a = 1") unless this has already been done, in which case it is a comparison operator (eg.
					// the second equals sign in "a = b = c", meaning compare "b" to "c" and set "a" to be the result of that)
					inExpressionContent = true;
					continue;
				}

				if (token is OpenBrace)
					bracketCount++;
				else if (token is CloseBrace)
				{
					if (bracketCount == 0)
						throw new ArgumentException("Invalid input, mismatched brackets");
					bracketCount--;
				}

				if (inExpressionContent)
					expressionTokens.Add(token);
				else
					valueToSetTokens.Add(token);
			}

			// If we encountered a "=" which caused a switch from value-to-set tokens to expression-to-set-value-to tokens then this
			// must be a ValueSettingStatement. Otherwise, it's a non-value-setting statement, all of the tokens for which will be
			// in the valueToSet set.
			if (inExpressionContent)
			{
				return new ValueSettingStatement(
					new Expression(valueToSetTokens),
					new Expression(expressionTokens),
					isSetStatement ? ValueSettingStatement.ValueSetTypeOptions.Set : ValueSettingStatement.ValueSetTypeOptions.Let
				);
			}
			return new Statement(
				valueToSetTokens,
				hasCallPrefix ? Statement.CallPrefixOptions.Present : Statement.CallPrefixOptions.Absent
			);
		}
	}
}
