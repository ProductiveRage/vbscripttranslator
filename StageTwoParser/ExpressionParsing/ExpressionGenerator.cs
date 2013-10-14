using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
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
            var expressionSegmentBuffer = new List<IExpressionSegment>();
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

                    expressionSegmentBuffer.AddRange(
                        GetExpressionSegments(
                            accessorBuffer,
                            Generate(argumentBuffer),
                            0
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
                    accessorBuffer.Add(token);
                    continue;
                }

                if (token is ArgumentSeparatorToken)
                {
                    if (accessorBuffer.Any())
                    {
                        expressionSegmentBuffer.AddRange(
                            GetExpressionSegments(
                                accessorBuffer,
                                new Expression[0],
                                0
                            )
                        );
                        accessorBuffer.Clear();
                    }
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
                expressionSegmentBuffer.AddRange(
                    GetExpressionSegments(
                        accessorBuffer,
                        new Expression[0],
                        0
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

        private static IEnumerable<IExpressionSegment> GetExpressionSegments(IEnumerable<IToken> accessorBuffer, IEnumerable<Expression> arguments, int callDepth)
        {
            if (accessorBuffer == null)
                throw new ArgumentNullException("accessorBuffer");
            if (arguments == null)
                throw new ArgumentNullException("arguments");
            if (callDepth < 0)
                throw new ArgumentOutOfRangeException("callDepth", "must be zero or greater");

            if (accessorBuffer.Any())
            {
                return SplitOnOperators(
                    accessorBuffer,
                    arguments,
                    callDepth
                );
            }
            if (!arguments.Any())
                throw new ArgumentException("No accessorBuffer content and no arguments content means that no expression segment can be generated");
            return new[]
            {
                new BracketedExpressionSegment(arguments)
            };
        }

        /// <summary>
        /// TODO: Expectations (no brackets, etc..)
        /// </summary>
        private static IEnumerable<IExpressionSegment> SplitOnOperators(IEnumerable<IToken> tokens, IEnumerable<Expression> arguments, int callDepth)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (arguments == null)
                throw new ArgumentNullException("arguments");
            if (callDepth < 0)
                throw new ArgumentOutOfRangeException("callDepth", "must be zero or greater");

            var tokensArray = tokens.ToArray();
            if (!tokensArray.Any())
                throw new ArgumentException("Empty tokens set - invalid");
            if (tokensArray.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in tokens set");
            if (tokensArray.Any(t => (t is OpenBrace) || (t is CloseBrace)))
                throw new ArgumentException("Open/CloseBrace encountered in tokens set - these may not be present in token sets at this point");

            // See http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx: "arithmetic operators are evaluated first, comparison operators are evaluated
            // next, and logical operators are evaluated last". The information from that article is also incorporated into the token values sets used below,
            // they appear in the lists in ordere of precedence.
            var breakOnOptions = new List<Predicate<IToken>>();
            foreach (var value in AtomToken.OperatorTokenValues)
                breakOnOptions.Add(token => (token is OperatorToken) && token.Content.Equals(value, StringComparison.InvariantCultureIgnoreCase));
            foreach (var value in AtomToken.ComparisonTokenValues)
                breakOnOptions.Add(token => (token is ComparisonToken) && token.Content.Equals(value, StringComparison.InvariantCultureIgnoreCase));
            foreach (var value in AtomToken.LogicalOperatorTokenValues)
                breakOnOptions.Add(token => (token is LogicalOperatorToken) && token.Content.Equals(value, StringComparison.InvariantCultureIgnoreCase));

            // Group by comparison first
            // - Order according to: =, <>, <, >, <=, >=, Is
            // Group by operator next
            // - Order according to: ^, /, *, \, Mod, +, -, &

            var breakOnIndex = -1;
            foreach (var breakOn in breakOnOptions)
            {
                for (var index = tokensArray.Length - 1; index >= 0; index--)
                {
                    if (breakOn(tokensArray[index]))
                    {
                        breakOnIndex = index;
                        break;
                    }
                }
            }
            if (breakOnIndex == -1)
            {
                return new[]
                {
                    new CallExpressionSegment(
                        tokensArray.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                        arguments
                    )
                };
            }

            var left = tokensArray.Take(breakOnIndex);
            var split = tokensArray[breakOnIndex];
            var right = tokensArray.Skip(breakOnIndex + 1);

            var coreSegments = GetExpressionSegments(left, new Expression[0], callDepth + 1)
                .Concat(new[] { new OperatorOrComparisonExpressionSegment(split) })
                .Concat(GetExpressionSegments(right, arguments, callDepth + 1));

            if (callDepth == 0)
            {
                // Only wrap the segments in brackets if we've drilled down through the content, we don't need to apply
                // them at the outer layer (otherwise unnecessary bracketing is being introduced)
                // - eg. "a + b + c" should become "(a + b) + c", not "((a + b) + c)"
                return coreSegments;
            }

            return new[]
            {
                new BracketedExpressionSegment(
                    new[]
                    {
                        new Expression(coreSegments)
                    }
                )
            };
       }
    }
}
