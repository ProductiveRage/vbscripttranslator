using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public static class ExpressionGenerator
    {
        /// <summary>
        /// This will arrange a token set into an expression tree where no expression will have more than one operator, where multiple operators are
        /// present the terms will be bracketed up to apply the max-one-operator restriction and to enforce VBScript operator precedence. This will
        /// never return null nor a set containing any nulls, it will raise an exception for a null token set or a set containing any nulls.
        /// </summary>
        public static IEnumerable<Expression> Generate(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            return Generate(new TokenNavigator(tokens), 0);
        }

        /// <summary>
        /// This will never return null nor a set containing any nulls
        /// </summary>
        private static IEnumerable<Expression> Generate(TokenNavigator tokenNavigator, int depth)
        {
            if (tokenNavigator == null)
                throw new ArgumentNullException("tokenNavigator");
            if (depth < 0)
                throw new ArgumentOutOfRangeException("depth", "must be zero or greater");

            var expressions = new List<Expression>();
            var expressionSegments = new List<IExpressionSegment>();
            var accessorBuffer = new List<IToken>();
            while (true)
            {
                var token = tokenNavigator.Value;
                if (token == null)
                {
                    if (depth > 0)
                        throw new ArgumentException("Expected CloseBrace and didn't encounter one - invalid content");
                    break;
                }

                if (token is CloseBrace)
                {
                    if (depth == 0)
                        throw new ArgumentException("Unexpected CloseBrace - invalid content");
                    tokenNavigator.MoveNext(); // Move on since this token has been processed
                    break;
                }

                if (token is ArgumentSeparatorToken)
                {
                    if (depth == 0)
                        throw new ArgumentException("Encountered ArgumentSeparatorToken in top-level content - invalid");

                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(accessorBuffer, new Expression[0])
                        );
                        accessorBuffer.Clear();
                    }
                    if (!expressionSegments.Any())
                        throw new ArgumentException("Unexpected ArgumentSeparatorToken - invalid content");

                    expressions.Add(GetExpression(expressionSegments));
                    expressionSegments.Clear();
                    tokenNavigator.MoveNext(); // Move on since this token has been processed
                    continue;
                }

                if (token is OpenBrace)
                {
                    tokenNavigator.MoveNext(); // Move on since this token has been processed

                    // Get the content from inside the brackets (using a TokenNavigator here that is passed again into the Generate
                    // method means that when the below call returns, the tokenNavigator here will have been moved along to after
                    // the bracketed content that is about to be processed)
                    var bracketedExpressions = Generate(tokenNavigator, depth + 1);

                    // If the accessorBuffer isn't empty then the bracketed content should be arguments, if not then it's just a bracketed expression
                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(
                                accessorBuffer,
                                bracketedExpressions
                            )
                        );
                        accessorBuffer.Clear();
                    }
                    else if (bracketedExpressions.Any())
                    {
						if (bracketedExpressions.Count() > 1)
							throw new ArgumentException("If bracketed content is not for an argument list then it's invalid for there to be multiple expressions within it");
						expressionSegments.Add(
							WrapExpressionSegments(bracketedExpressions.Single().Segments)
                        );
                    }
                    continue;
                }

                var operatorToken = token as OperatorToken;
                if (operatorToken != null)
                {
                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(accessorBuffer, new Expression[0])
                        );
                        accessorBuffer.Clear();
                    }
                    expressionSegments.Add(
                        new OperationExpressionSegment(operatorToken)
                    );
                    tokenNavigator.MoveNext();
                    continue;
                }

                accessorBuffer.Add(token);
                tokenNavigator.MoveNext();
            }
            if (accessorBuffer.Any())
            {
                expressionSegments.Add(
                    GetCallOrNewOrValueExpressionSegment(accessorBuffer, new Expression[0])
                );
                accessorBuffer.Clear();
            }
            if (expressionSegments.Any())
            {
                expressions.Add(GetExpression(expressionSegments));
                expressionSegments.Clear();
            }
            return expressions;
        }

        /// <summary>
        /// This applies bracketing to expression segments to enforce VBScript's operator precedence rules
        /// </summary>
        private static Expression GetExpression(IEnumerable<IExpressionSegment> segments)
        {
            if (segments == null)
                throw new ArgumentNullException("segments");

            var segmentsArray = segments.ToArray();
            if (!segmentsArray.Any())
                throw new ArgumentException("Empty segments set specified - invalid");
            if (segmentsArray.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in segments set");

            // If there are zero or one operators in the expression content then we don't need to do any bracketing to guarantee operator precedence
            var operatorSegments = segmentsArray
                .Select((s, index) => Tuple.Create(s as OperationExpressionSegment, index))
                .Where(s => s.Item1 != null);
            if (operatorSegments.Count() < 2)
                return GetCallExpressionSegmentGroupedExpression(segmentsArray);

            // See http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx: "arithmetic operators are evaluated first, comparison operators are evaluated
            // next, and logical operators are evaluated last". The information from that article is also incorporated into the AtomToken ComparisonTokenValues,
            // ComparisonTokenValues and ArithmeticAndStringOperatorTokenValues sets. Where there are multiple operations with the same precedence, they are
            // processed from left-to-right.
            // Note: There are two operations which only have a right-hand-side - arithmetic negation and logical inversion. The first is always applied
            // first but has the complication that it must not be confused with a subtraction (an operation with the same operation token but with both
            // a left and hand side). The NOT operator is the first of the logical operators to be applied but logical operations are only addressed
            // after arithmetic (and string concatenation) and comparisons.
            
            // First, check for an arithmetic and wrap that operation and the corresponding value, feeding back into this method (this is the first special case)
            var firstArithmeticNegationOperation = operatorSegments.FirstOrDefault(s =>
                (s.Item1.Token.Content == "-") && ((s.Item2 == 0) || (segmentsArray[s.Item2 - 1] is OperationExpressionSegment))
            );
            if (firstArithmeticNegationOperation != null)
            {
                if (firstArithmeticNegationOperation.Item2 == (segmentsArray.Length - 1))
                    throw new ArgumentException("Encountered negation operator at end of expression segment content - invalid");
                return GetExpression(
                    BracketOffTerms(segmentsArray, firstArithmeticNegationOperation.Item2, 2)
                );
            }

            // Next, check whether any NOT handling is required - this is only the case if the only operations are logical (boolean) operations
            if (operatorSegments.All(s => s.Item1.Token is LogicalOperatorToken))
            {
                var firstLogicalInversion = operatorSegments.FirstOrDefault(s =>
                    s.Item1.Token.Content.Equals("NOT", StringComparison.InvariantCultureIgnoreCase)
                );
                if (firstLogicalInversion != null)
                {
                    if (firstLogicalInversion.Item2 == (segmentsArray.Length - 1))
                        throw new ArgumentException("Encountered NOT operator at end of expression segment content - invalid");
                    return GetExpression(
                        BracketOffTerms(segmentsArray, firstLogicalInversion.Item2, 2)
                    );
                }
            }

            // Now that the special cases are out of the way, we just need to determine which of the current operators needs addressing first
            // - We follow the VBScript rules of precedence and then reverse them in order to determine where to break on (if we have the expression
            //   "a + b * c" then "b * c" should be bracketed since multiplication takes precedence, so we break on the "+" and bracket off the b,
            //   c multiplication against the remaining a token - so we had to break on the operator with the least precedence, hence taking the
            //   last element of the ordered set)
            var segmentToBreakOn = operatorSegments
                .OrderBy(s => s, new IndexerOperationExpressionSegmentSorter())
                .Last();

            var left = segmentsArray.Take(segmentToBreakOn.Item2);
            var right = segmentsArray.Skip(segmentToBreakOn.Item2 + 1);
            return GetCallExpressionSegmentGroupedExpression(new IExpressionSegment[]
            {
                WrapExpressionSegments(GetExpression(left).Segments),
                segmentToBreakOn.Item1,
                WrapExpressionSegments(GetExpression(right).Segments)
            });
        }

        private static IEnumerable<IExpressionSegment> BracketOffTerms(IEnumerable<IExpressionSegment> segments, int index, int count)
        {
            if (segments == null)
                throw new ArgumentNullException("segments");
            if (index < 0)
                throw new ArgumentOutOfRangeException("index", "must be zero or greater");
            if (count < 2)
                throw new ArgumentOutOfRangeException("count", "must be at least two"); // Otherwise there's no point bracketing off the segments!
            
            var segmentsArray = segments.ToArray();
            if (segmentsArray.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in segments set");
            if (index > (segmentsArray.Length - 1))
                throw new ArgumentOutOfRangeException("index", "must refer to a location in the segments set");
            if ((index + count) > segmentsArray.Length)
                throw new ArgumentOutOfRangeException("count", "index + count must not go beyond the end of the segments set");

            return segmentsArray.Take(index)
                .Concat(new[] {
                    WrapExpressionSegments(
                        segmentsArray.Skip(index).Take(count)
                    )
                })
                .Concat(
                    segmentsArray.Skip(index + count)
                );
        }

        /// <summary>
        /// If a single token of type NumericValueToken or StringToken is specified then return a NumericValueExpressionSegment or StringValueExpressionSegment to
        /// represent the constant value (there should be zero arguments in this case). If there are two tokens, a KeyWordToken with content "new" and NameToken
        /// (with no arguments) then this can be represented by a NewInstanceExpressionSegment. Otherwise return a CallExpressionSegment.
        /// that data in a constant-type expression segment).
        /// </summary>
        private static IExpressionSegment GetCallOrNewOrValueExpressionSegment(IEnumerable<IToken> tokens, IEnumerable<Expression> arguments)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var tokensArray = tokens.ToArray();
            if (!tokensArray.Any())
                throw new ArgumentException("Empty tokens set specified, invalid");
            if (tokensArray.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in tokens set");

            // If there are arguments then there's no change of representing this as a constant-type expression or as a new instance request
            if (!arguments.Any())
            {
                if (tokensArray.Length == 1)
                {
                    var numericValue = tokensArray[0] as NumericValueToken;
                    if (numericValue != null)
                        return new NumericValueExpressionSegment(numericValue);
                    var stringValue = tokensArray[0] as StringToken;
                    if (stringValue != null)
                        return new StringValueExpressionSegment(stringValue);
                }
                else if ((tokensArray.Length == 2)
                && (tokensArray[0] is KeyWordToken)
                && tokensArray[0].Content.Equals("new", StringComparison.InvariantCultureIgnoreCase))
                {
                    var newInstanceName = tokensArray[1] as NameToken;
                    if (newInstanceName != null)
                        return new NewInstanceExpressionSegment(newInstanceName);
                }
            }
            return new CallExpressionSegment(
                tokensArray.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                arguments
            );
        }

        /// <summary>
		/// Generate a BracketedExpressionSegment instance - if there is only a single expression segment, where that segment is a bracketed segment, then return
		/// that segment rather than wrapping it again (this is done recursively in case there are multiple layers of over-wrapped bracketed segments). Note: If
		/// it ends up that there's only one expression segment in total then this will be returned, unwrapped.
		/// </summary>
		private static IExpressionSegment WrapExpressionSegments(IEnumerable<IExpressionSegment> segments)
		{
			if (segments == null)
				throw new ArgumentNullException("segments");

			var segmentsArray = segments.ToArray();
			if (segmentsArray.Any(s => s == null))
				throw new ArgumentException("Null reference encountered in segments set");

			while (true)
			{
				if (segmentsArray.Length != 1)
					break;

				var onlySegmentAsBracketedSegment = segmentsArray[0] as BracketedExpressionSegment;
				if (onlySegmentAsBracketedSegment == null)
					break;

				segmentsArray = onlySegmentAsBracketedSegment.Segments.ToArray();
			}
			if (segmentsArray.Length == 1)
			{
				// If there's only one term to wrap then we can just return that without any wrapping!
				return segmentsArray[0];
			}
			return new BracketedExpressionSegment(segmentsArray);
		}

        /// <summary>
        /// Consecutive CallExpressionSegments should not appear in an Expression as consecutive CallExpressionSegments are segments that represent part
        /// of the same operation - eg. "a(0).Test" is represented by two CallExpressionSegments, one for "a(0)" and one for "Test" - but really they
        /// describe two parts of a single retrieval. As such, they should be wrapped in a CallSetExpressionSegment (single CallExpressionSegments
        /// are fine).
        /// </summary>
        private static Expression GetCallExpressionSegmentGroupedExpression(IEnumerable<IExpressionSegment> segments)
        {
            if (segments == null)
                throw new ArgumentNullException("segments");

            var callExpressionSegmentBuffer = new List<CallExpressionSegment>();
            var expressionSegments = new List<IExpressionSegment>();
            foreach (var segment in segments)
            {
                if (segment == null)
                    throw new ArgumentException("Null reference encountered in segments set");

                var callExpressionSegment = segment as CallExpressionSegment;
                if (callExpressionSegment != null)
                {
                    callExpressionSegmentBuffer.Add(callExpressionSegment);
                    continue;
                }

                if (callExpressionSegmentBuffer.Count > 0)
                {
                    if (callExpressionSegmentBuffer.Count == 1)
                        expressionSegments.Add(callExpressionSegmentBuffer[0]);
                    else
                        expressionSegments.Add(new CallSetExpressionSegment(callExpressionSegmentBuffer));
                    callExpressionSegmentBuffer.Clear();
                }
                expressionSegments.Add(segment);
            }
            if (callExpressionSegmentBuffer.Count > 0)
            {
                if (callExpressionSegmentBuffer.Count == 1)
                    expressionSegments.Add(callExpressionSegmentBuffer[0]);
                else
                    expressionSegments.Add(new CallSetExpressionSegment(callExpressionSegmentBuffer));
            }
            return new Expression(expressionSegments);
        }

        private static NewInstanceExpressionSegment TryToGetAsNewInstanceExpression(IExpressionSegment segment)
        {
            if (segment == null)
                throw new ArgumentNullException("segment");

            CallExpressionSegment callExpressionSegment;
            if (segment is CallSetExpressionSegment)
            {
                var callSetExpressionSegment = segment as CallSetExpressionSegment;
                if (callSetExpressionSegment.CallExpressionSegments.Count() > 1)
                    return null;
                callExpressionSegment = callSetExpressionSegment.CallExpressionSegments.Single();
            }
            else
                callExpressionSegment = segment as CallExpressionSegment;
            if (callExpressionSegment == null)
                return null;

            if ((callExpressionSegment.MemberAccessTokens.Count() != 2) || callExpressionSegment.Arguments.Any())
                return null;

            var memberAccessToken0 = callExpressionSegment.MemberAccessTokens.First();
            if (!memberAccessToken0.Content.Equals("new", StringComparison.InvariantCultureIgnoreCase))
                return null;

            var memberAccessToken1 = callExpressionSegment.MemberAccessTokens.ElementAt(1) as NameToken;
            if (memberAccessToken1 == null)
                return null;

            return new NewInstanceExpressionSegment(memberAccessToken1);
        }

        private class IndexerOperationExpressionSegmentSorter : IComparer<Tuple<OperationExpressionSegment, int>>
        {
            public int Compare(Tuple<OperationExpressionSegment, int> x, Tuple<OperationExpressionSegment, int> y)
            {
                if (x == null)
                    throw new ArgumentNullException("x");
                if (y == null)
                    throw new ArgumentNullException("y");

                var operatorTypeValueX = GetOperatorTypeValue(x.Item1);
                var operatorTypeValueY = GetOperatorTypeValue(y.Item1);
                if (operatorTypeValueX != operatorTypeValueY)
                    return operatorTypeValueX.CompareTo(operatorTypeValueY);

                var contentValueX = GetContentValue(x.Item1);
                var contentValueY = GetContentValue(y.Item1);
                if (contentValueX != contentValueY)
                    return contentValueX.CompareTo(contentValueY);

                return x.Item2.CompareTo(y.Item2);
            }

            private static IEnumerable<string> AllOperatorContentStrings =
                AtomToken.ArithmeticAndStringOperatorTokenValues
                .Concat(AtomToken.ComparisonTokenValues)
                .Concat(AtomToken.LogicalOperatorTokenValues);

            private static int GetOperatorTypeValue(OperationExpressionSegment segment)
            {
                if (segment == null)
                    throw new ArgumentNullException("segment");

                if (segment.Token is LogicalOperatorToken)
                    return 3;
                else if (segment.Token is ComparisonOperatorToken)
                    return 2;
                else
                    return 1;
            }

            private static int GetContentValue(OperationExpressionSegment segment)
            {
                if (segment == null)
                    throw new ArgumentNullException("segment");

                var operatorContentOptions =
                    AtomToken.ArithmeticAndStringOperatorTokenValues
                    .Concat(AtomToken.ComparisonTokenValues)
                    .Concat(AtomToken.LogicalOperatorTokenValues)
                    .Select((value, index) => new { Value = value, Index = index })
                    .FirstOrDefault(c => c.Value.Equals(segment.Token.Content, StringComparison.InvariantCultureIgnoreCase));
                if (operatorContentOptions == null)
                    throw new NotSupportedException("Unrecognised operator token value");
                return operatorContentOptions.Index;
            }
        }

        private class TokenNavigator
        {
            private readonly IToken[] _tokens;
            private int _index;
            public TokenNavigator(IEnumerable<IToken> tokens)
            {
                if (tokens == null)
                    throw new ArgumentNullException("tokens");

                _tokens = tokens.ToArray();
                if (_tokens.Any(t => t == null))
                    throw new ArgumentException("Null reference encountered in tokens set");
                _index = 0;
            }

            public IToken Value
            {
                get { return PastEndOfContent ? null : _tokens[_index]; }
            }

            public void MoveNext()
            {
                if (!PastEndOfContent)
                    _index++;
            }

            private bool PastEndOfContent
            {
                get { return _index >= _tokens.Length; }
            }
        }
    }
}
