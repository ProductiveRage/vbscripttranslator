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
    /// as the negative signs cancel each other out. There may still be some unnecessary "+" symbols as "1 / +2" is valid in VBScript, which may be
    /// reduced to "1 / 2". Note that "1 / -2" would not be altered (but the NumberRebuilder should be used to construct a NumericValueToken which
    /// incorporates both the "-" and the "2", leaving only a single operator; the "/"). This method is also responsible for combinination of
    /// particular operators such as ">" and "=" becoming "=". The tokens must be passed through the NumberRebuilder before applying this
    /// processing.
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
            var previousTokenIfAny = (IToken)null;
            foreach (var token in tokens)
            {
                if (token == null)
                    throw new ArgumentException("Null reference encountered in tokens set");

                var combinableOperator = TryToGetAsAdditionOrSubtractionToken(token);
                if (combinableOperator == null)
                {
                    var bufferHadContentThatWasReducedToNothing = false;
                    if (buffer.Any())
                    {
                        var condensedToken = CondenseNegations(buffer);
                        if (IsTokenRedundant(condensedToken, previousTokenIfAny))
                        {
                            // If this is a "+" and the last token was an OperatorToken, then this one is redundant (eg. "1 * +1")
                            bufferHadContentThatWasReducedToNothing = true;
                        }
                        else
                            additionSubtractionRewrittenTokens.Add(condensedToken);
                        buffer.Clear();
                    }

                    // When a minus-sign/addition-sign buffer is flattened and can be reduced to nothing, if the next token is a numeric value then we
                    // need to apply a bit of a dirty hack since VBScript gives numeric literals special treatment in some cases but does not consider
                    // --1 to be a numeric literal (for example). So we can not replace --1 with 1 since it would change the meaning of some code. To
                    // illustrate, consider the following:
                    //   If ("a" = 1) Then
                    //   If ("a" = --1) Then
                    //   If ("a" = +-1) Then
                    // The first example will result in a Type Mismatch since the numeric literal forces the "a" to be parsed as a number (which fails).
                    // However, the second and third examples return false since their right hand side values are not considered to be numeric literals
                    // and so the left hand sides need not be parsed as numeric values. The workaround is to identify these situations and to wrap the
                    // number if a CDbl call. This will not affect the numeric value but it will prevent it from being identified as a numeric literal
                    // later on (this is important to the StatementTranslator).Note: This is why the NumberRebuilder must have done its work before
                    // we get here, since ++1.2 must be recognised as "+", "+", "1.2" so that it can be translated into "CDbl(1.2)", rather than
                    // still being "+", "+", "1", ".", "2", which would translated into "CDbl(1).2", which would be invalid.
                    var wrapTokenInCDblCall = bufferHadContentThatWasReducedToNothing && (token is NumericValueToken);
                    if (wrapTokenInCDblCall)
                    {
                        additionSubtractionRewrittenTokens.Add(new BuiltInFunctionToken("CDbl", token.LineIndex));
                        additionSubtractionRewrittenTokens.Add(new OpenBrace(token.LineIndex));
                    }
                    additionSubtractionRewrittenTokens.Add(token);
                    if (wrapTokenInCDblCall)
                        additionSubtractionRewrittenTokens.Add(new CloseBrace(token.LineIndex));
                    previousTokenIfAny = token;
                }
                else
                    buffer.Add(combinableOperator);
            }
            if (buffer.Any())
            {
                // Note: We don't need to copy all of the logic from above - in fact we can't, since we don't have a current token reference
                var condensedToken = CondenseNegations(buffer);
                if (!IsTokenRedundant(condensedToken, previousTokenIfAny))
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
                    comparisonRewrittenTokens.Add(AtomToken.GetNewToken(token.Content + nextToken.Content, token.LineIndex));
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
            if (!tokens.Any())
                throw new ArgumentException("Empty tokens set specified - invalid");

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
            return new OperatorToken(isNegative ? "-" : "+", tokens.First().LineIndex);
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
