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
        /// This will never return null nor a set containing any nulls
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
                            new CallExpressionSegment(
                                accessorBuffer.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                                new Expression[0]
                            )
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
                            new CallExpressionSegment(
                                accessorBuffer.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                                bracketedExpressions
                            )
                        );
                        accessorBuffer.Clear();
                    }
                    else
                    {
                        expressionSegments.Add(
                            WrapExpressionSegments(bracketedExpressions)
                        );
                    }
                    continue;
                }

                if (token is OperatorToken)
                {
                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            new CallExpressionSegment(
                                accessorBuffer.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                                new Expression[0]
                            )
                        );
                        accessorBuffer.Clear();
                    }
                    expressionSegments.Add(
                        new OperationExpressionSegment(token)
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
                    new CallExpressionSegment(
                        accessorBuffer.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                        new Expression[0]
                    )
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
                return new Expression(segmentsArray);

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
            return new Expression(new IExpressionSegment[]
            {
                WrapExpressionSegments(new[] { GetExpression(left) }),
                segmentToBreakOn.Item1,
                WrapExpressionSegments(new[] { GetExpression(right) })
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
                    WrapExpressionSegments(new[]
                    {
                        new Expression(
                            segmentsArray.Skip(index).Take(count)
                        )
                    })
                })
                .Concat(
                    segmentsArray.Skip(index + count)
                );
        }

        /// <summary>
		/// Generate a BracketedExpressionSegment instance - if there is only a single expression specified, with a single expression segment, where that segment
		/// is a bracketed segment, then return that segment rather than wrapping it again (this is done recursively in case there are multiple layers of over-
		/// wrapped bracketed segments). Note: If it ends up that there's only one expression segment in total then this will be returned, unwrapped.
		/// </summary>
		private static IExpressionSegment WrapExpressionSegments(IEnumerable<Expression> expressions)
		{
			if (expressions == null)
				throw new ArgumentNullException("segments");

			var expressionsArray = expressions.ToArray();
			if (expressionsArray.Any(e => e == null))
				throw new ArgumentException("Null reference encountered in expressions set");

			while (true)
			{
				if (expressionsArray.Length != 1)
					break;

				var onlyExpression = expressionsArray[0];
				if (onlyExpression.Segments.Count() != 1)
					break;

				var onlySegmentAsBracketedSegment = onlyExpression.Segments.Single() as BracketedExpressionSegment;
				if (onlySegmentAsBracketedSegment == null)
					break;

				if (onlySegmentAsBracketedSegment.Expressions.Count() != 1)
					break;

				expressionsArray = new[] { onlySegmentAsBracketedSegment.Expressions.Single() };
			}
            if ((expressionsArray.Length == 1) && (expressionsArray[0].Segments.Count() == 1))
            {
                // If there's only one term to wrap then we can just return that without any wrapping!
                return expressionsArray[0].Segments.Single();
            }
			return new BracketedExpressionSegment(expressionsArray);
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
