using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser
{
    public static class ExpressionGenerator
    {
        /// <summary>
        /// This will never return null nor a set containing any nulls
        /// </summary>
        public static IEnumerable<Expression> Generate(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var expressions = new List<Expression>();
            var expressionSegmentBuffer = new List<ExpressionSegment>();
            var tokenArray = tokens.ToArray();
            var accessorBuffer = new List<IToken>();
            for (var index = 0; index < tokenArray.Length; index++)
            {
                var token = tokenArray[index];
                if (token == null)
                    throw new ArgumentException("Null reference encountered in tokens set");

                if (token is OpenBrace)
                {
                    var bracketCount = 1;
                    var argumentBuffer = new List<IToken>();
                    while (index < tokenArray.Length)
                    {
                        index++;
                        token = tokenArray[index];
                        if (token is OpenBrace)
                            bracketCount++;
                        else if (token is CloseBrace)
                        {
                            bracketCount--;
                            if (bracketCount == 0)
                                break;
                        }
                        argumentBuffer.Add(token);
                    }
                    if (bracketCount != 0)
                        throw new ArgumentException("Invalid content - mismatched brackets");

                    expressionSegmentBuffer.Add(
                        new ExpressionSegment(
                            accessorBuffer,
                            Generate(argumentBuffer)
                        )
                    );
                    accessorBuffer.Clear();
                    argumentBuffer.Clear();
                    continue;
                }

                if (token is MemberAccessorOrDecimalPointToken)
                {
                    if (!(token is MemberAccessorToken))
                    {
                        // There shouldn't be any "." character that may be a member accessor or a decimal point, work to differentiate
                        // should have been before this point. The NumberRebuilder class can be used for this.
                        throw new ArgumentException("Encountered a MemberAccessorOrDecimalPointToken, these should all have been removed or specialised as MemberAccessorToken instances before calling this");
                    }
                    /* TODO: Confirm behaviour
                    if (accessorBuffer.Any())
                    {
                        expressionSegmentBuffer.Add(
                            new ExpressionSegment(
                                accessorBuffer,
                                new Expression[0]
                            )
                        );
                        accessorBuffer.Clear();
                    }
                    if (!expressionSegmentBuffer.Any())
                        throw new ArgumentException("Invalid content - orphan member access token (\"" + token.Content + "\")");
                     */
                    // TODO: Explain why not adding
                    //accessorBuffer.Add(token);
                    continue;
                }

                if (token is ArgumentSeparatorToken)
                {
                    if (!expressionSegmentBuffer.Any())
                        throw new ArgumentException("Invalid content - unexpected argument separator access token (\"" + token.Content + "\")");
                    expressions.Add(
                        new Expression(expressionSegmentBuffer)
                    );
                    expressionSegmentBuffer.Clear();
                    continue;
                }

                accessorBuffer.Add(token);
            }
            if (accessorBuffer.Any())
            {
                expressionSegmentBuffer.Add(
                    new ExpressionSegment(
                        accessorBuffer,
                        new Expression[0]
                    )
                );
            }
            if (expressionSegmentBuffer.Any())
            {
                expressions.Add(
                    new Expression(expressionSegmentBuffer)
                );
            }
            return expressions;
        }
    }
}
