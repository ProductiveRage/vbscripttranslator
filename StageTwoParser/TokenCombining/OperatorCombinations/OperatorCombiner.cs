using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations
{
    /// <summary>
    /// In VBScript, arbitrary runs of + and - operators are acceptable (this is not the case for any other operators). These can be condensed without
    /// changing the meaning. The most obvious is "1 + -1" which may become "1 - 1" but "1-+---+2" is allowable in VBScript but may be reduced to "1+2"
    /// as the negative signs cancel each other out. There may still be some unnecessary "+" symbols are "1 / +2" is valid in VBScript, which may be
    /// reduced to "1 / 2". Note that "1 / -2" would not be altered (but the NumberRebuilder should be used to construct a NumericValueToken which
    /// incorporates both the "-" and the "2", leaving only a single operator; the "/"). This method is also responsible for combinination of
    /// particular operators such as ">" and "=" becoming "=".
    /// </summary>
    public static class OperatorCombiner
    {
        public static IEnumerable<IToken> Combine(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            // Handle +/- sign combinations
            var additionSubtractionRewrittenTokens = new List<IToken>();
            var buffer = new List<OperatorToken>();
            var previousToken = (IToken)null;
            foreach (var token in tokens)
            {
                if (token == null)
                    throw new ArgumentException("Null reference encountered in tokens set");

                var combinableOperator = TryToGetAsAdditionOrSubtractionToken(token);
                if (combinableOperator == null)
                {
                    if (buffer.Any())
                    {
                        var condensedToken = CondenseNegations(buffer);
                        if (!IsTokenRedundant(condensedToken, previousToken))
                        {
                            // If this is a "+" and the last token was an OperatorToken, then this one is redundant (eg. "1 * +1")
                            additionSubtractionRewrittenTokens.Add(condensedToken);
                        }
                        buffer.Clear();
                    }
                    additionSubtractionRewrittenTokens.Add(token);
                    previousToken = token;
                }
                else
                    buffer.Add(combinableOperator);
            }
            if (buffer.Any())
            {
                var condensedToken = CondenseNegations(buffer);
                if (!IsTokenRedundant(condensedToken, previousToken))
                    additionSubtractionRewrittenTokens.Add(condensedToken);
            }

            // Handle comparison token combinations (eg. ">", "=" to ">=")
            var combinations = new[]
            {
                Tuple.Create(Tuple.Create("<", ">"), "<>"),
                Tuple.Create(Tuple.Create("<", "="), "<="),
                Tuple.Create(Tuple.Create(">", "="), ">=")
            };
            var comparisonRewrittenTokens = new List<IToken>();
            for (var index = 0; index < additionSubtractionRewrittenTokens.Count; index++)
            {
                var token = additionSubtractionRewrittenTokens[index];
                if (index == (additionSubtractionRewrittenTokens.Count - 1))
                {
                    comparisonRewrittenTokens.Add(token);
                    continue;
                }

                var nextToken = additionSubtractionRewrittenTokens[index + 1];
                var combineTokens = (
                    ((token.Content == "<") && (nextToken.Content == ">")) ||
                    ((token.Content == ">") && (nextToken.Content == "=")) ||
                    ((token.Content == "<") && (nextToken.Content == "="))
                );
                if (combineTokens)
                {
                    comparisonRewrittenTokens.Add(AtomToken.GetNewToken(token.Content + nextToken.Content));
                    index++;
                    continue;
                }
                comparisonRewrittenTokens.Add(token);
            }
            return comparisonRewrittenTokens;
        }

        /// <summary>
        /// This will return null for any token that is not an addition or subtraction OperatorToken (a null token will result in an exception being raised)
        /// </summary>
        private static OperatorToken TryToGetAsAdditionOrSubtractionToken(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            var operatorToken = token as OperatorToken;
            if (operatorToken == null)
                return null;

            return ((operatorToken.Content == "+") || (operatorToken.Content == "-")) ? operatorToken : null;
        }

        /// <summary>
        /// These will reduce streams of addition and/or subtraction OperationTokens to the minimum representation that won't affect the meaning
        /// </summary>
        private static OperatorToken CondenseNegations(IEnumerable<OperatorToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var isNegative = false;
            foreach (var token in tokens)
            {
                if (token == null)
                    throw new ArgumentException("Null reference encountered in tokens set");

                if (token.Content == "-")
                    isNegative = !isNegative;
                else if (token.Content != "+")
                    throw new ArgumentException("All tokens must be operators with content either \"-\" or \"+\"");
            }
            return new OperatorToken(isNegative ? "-" : "+");
        }

        private static bool IsTokenRedundant(IToken token, IToken previousTokenIfAny)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return (
                (token is OperatorToken) &&
                (token.Content == "+") &&
                (previousTokenIfAny != null) &&
                (previousTokenIfAny is OperatorToken)
            );
        }
    }
}
