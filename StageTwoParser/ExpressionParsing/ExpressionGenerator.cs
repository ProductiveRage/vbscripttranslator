using CSharpSupport.Exceptions;
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
        public static IEnumerable<Expression> Generate(IEnumerable<IToken> tokens, IToken directedWithReferenceIfAny, Action<string> warningLogger)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

            return Generate(new TokenNavigator(tokens), 0, directedWithReferenceIfAny, warningLogger);
        }

        /// <summary>
        /// This will never return null nor a set containing any nulls
        /// </summary>
        private static IEnumerable<Expression> Generate(TokenNavigator tokenNavigator, int depth, IToken directedWithReferenceIfAny, Action<string> warningLogger)
        {
            if (tokenNavigator == null)
                throw new ArgumentNullException("tokenNavigator");
            if (depth < 0)
                throw new ArgumentOutOfRangeException("depth", "must be zero or greater");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

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
                    if (expressionSegments.Count == 1)
                    {
                        var callExpressionSegment = expressionSegments[0] as CallExpressionSegment;
                        if ((callExpressionSegment != null) && (callExpressionSegment.MemberAccessTokens.Count() == 1) && !callExpressionSegment.Arguments.Any())
                        {
                            // VBScript gives special meaning to a reference wrapped in brackets when passing it to a function (or property); if
                            // the argument would otherwise be passed ByRef, it will be passed ByVal. If this close bracket terminates a section
                            // containing only a single CallExpressionSegment without property accesses or arguments, then the brackets have
                            // significance and the content must be represented by a BracketedExpressionSegment. (If it's a variable access with
                            // function / property accesses - meaning there would be multiple MemberAccessTokens - or if there are arguments -
                            // meaning that it's a function or property call - then the value would not be elligible for being passed ByRef anyway,
                            // so the extra brackets need not be maintained). There is chance that ths single token represents a non-argument call
                            // to a function, but we don't have enough information at this point to know that, so we have to presume the worst and
                            // keep the brackets here.
                            expressionSegments[0] = new BracketedExpressionSegment(new[] { callExpressionSegment });
                        }
                    }
                    break;
                }

                if (token is ArgumentSeparatorToken)
                {
                    if (depth == 0)
                        throw new ArgumentException("Encountered ArgumentSeparatorToken in top-level content - invalid");

                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(
								accessorBuffer,
								new Expression[0],
                                directedWithReferenceIfAny,
                                argumentsAreBracketed: false,
                                willBeFirstSegmentInCallExpression: WillBeFirstSegmentInCallExpression(expressionSegments),
                                warningLogger: warningLogger
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
                    var bracketedExpressions = Generate(tokenNavigator, depth + 1, directedWithReferenceIfAny, warningLogger);

                    // If the accessorBuffer isn't empty then the bracketed content should be arguments, if not then it's just a bracketed expression
                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(
                                accessorBuffer,
                                bracketedExpressions,
                                directedWithReferenceIfAny,
                                argumentsAreBracketed: true,
                                willBeFirstSegmentInCallExpression: WillBeFirstSegmentInCallExpression(expressionSegments),
                                warningLogger: warningLogger
                            )
                        );
                        accessorBuffer.Clear();
                    }
                    else if (bracketedExpressions.Any())
                    {
                        if (expressionSegments.Any() && (expressionSegments.Last() is CallSetItemExpressionSegment))
                        {
                            // If the previous expression segment was a CallExpressionSegment or CallSetItemExpressionSegment (the first
                            // is derived from the second so only a single type check is required) then this bracketed content should
                            // be considered a continuation of the call (these segments will later be grouped into a single
                            // CallSetExpression)
                            expressionSegments.Add(
                                new CallSetItemExpressionSegment(
                                    new IToken[0],
                                    bracketedExpressions,
                                    null // ArgumentBracketPresenceOptions
                                )
                            );
                        }
                        else
                        {
                            if (bracketedExpressions.Count() > 1)
                                throw new ArgumentException("If bracketed content is not for an argument list then it's invalid for there to be multiple expressions within it");
                            expressionSegments.Add(
                                WrapExpressionSegments(bracketedExpressions.Single().Segments, unwrapSingleBracketedTerm: false)
                            );
                        }
                    }
                    continue;
                }

                var operatorToken = token as OperatorToken;
                if (operatorToken != null)
                {
                    if (accessorBuffer.Any())
                    {
                        expressionSegments.Add(
                            GetCallOrNewOrValueExpressionSegment(
								accessorBuffer,
								new Expression[0],
                                directedWithReferenceIfAny,
                                argumentsAreBracketed: false, // zero-argument content not bracketed
                                willBeFirstSegmentInCallExpression: WillBeFirstSegmentInCallExpression(expressionSegments),
                                warningLogger: warningLogger
                            )
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
                    GetCallOrNewOrValueExpressionSegment(
						accessorBuffer,
						new Expression[0],
                        directedWithReferenceIfAny,
                        argumentsAreBracketed: false, // zero-argument content not bracketed
                        willBeFirstSegmentInCallExpression: WillBeFirstSegmentInCallExpression(expressionSegments),
                        warningLogger: warningLogger
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

        private static bool WillBeFirstSegmentInCallExpression(IEnumerable<IExpressionSegment> expressionSegmentBuffer)
        {
            if (expressionSegmentBuffer == null)
                throw new ArgumentNullException("expressionSegmentBuffer");

            // If the expressionSegmentBuffer is empty and the next segment to be to processed is a part of a call expression then it will definitely be the
            // first of its segments. Same logic applies if the previous expression segment was an operator since this will be the first part of an expression
            // that contains multiple expressions - eg. "a * b". Otherwise, the next expression segment must be part of a single expression that is being
            // processed (eg. ".Name" from "a.Name").
            var lastExpressionSegmentIfAny = expressionSegmentBuffer.LastOrDefault();
            return (lastExpressionSegmentIfAny == null) || (lastExpressionSegmentIfAny is OperationExpressionSegment);
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
            {
                if (operatorSegments.Any())
                {
                    // While it's true that we don't need to apply any more wrapping to expressions if there is only a single operator, we do want to UNWRAP any
                    // unnecessarily-wrapped segments, which the WrapExpressionSegments function achieves (when unwrapSingleBracketedTerm is true). This deals
                    // with cases such as
                    //   If (a = (1)) Then
                    // where the unnecessary brackets around the "1" need to be removed so that the comparison forces "a" into a number. This only needs to be
                    // done when there is an operator in the expression and only CAN be done in there is an operator, so that the brackets are not deemed
                    // unnecessary in
                    //   a = Test((b))
                    // since the "extra" brackets around "b" are NOT extraneous and indicate that "b" must be passed ByVal into Test.
                    return GetCallExpressionSegmentGroupedExpression(
                        segmentsArray.Select(s => WrapExpressionSegments(new[] { s }, unwrapSingleBracketedTerm: true))
                    );
                }
                return GetCallExpressionSegmentGroupedExpression(segmentsArray);
            }

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
            // - This is to handle cases such as "a AND NOT b" where "NOT b" must be combined and then considered by the "a AND x" operation, in
            //   other cases, "NOT" takes lower precedence (eg. "NOT a IS Nothing" is dealt with by combining "a IS Nothing" and then applying
            //   the NOT operation, this is dealt with further down)
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
            var operatorSegmentToBreakOn = operatorSegments
                .OrderBy(s => s, new IndexerOperationExpressionSegmentSorter())
                .Last();

            var left = segmentsArray.Take(operatorSegmentToBreakOn.Item2);
            var right = segmentsArray.Skip(operatorSegmentToBreakOn.Item2 + 1);
            var expressionSegmentsToGroup = new List<IExpressionSegment>();
            if (left.Any())
                expressionSegmentsToGroup.Add(WrapExpressionSegments(GetExpression(left).Segments, unwrapSingleBracketedTerm: true));
            else if (!operatorSegmentToBreakOn.Item1.Token.Content.Equals("NOT", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The content to the left of an operator may only be empty if it is a \"NOT\" logical operator");
            expressionSegmentsToGroup.Add(operatorSegmentToBreakOn.Item1);
            if (!right.Any())
                throw new ArgumentException("The content to the right of an operator may not be empty");
            expressionSegmentsToGroup.Add(WrapExpressionSegments(GetExpression(right).Segments, unwrapSingleBracketedTerm: true));
            return GetCallExpressionSegmentGroupedExpression(
                expressionSegmentsToGroup
            );
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
                    WrapExpressionSegments(segmentsArray.Skip(index).Take(count), unwrapSingleBracketedTerm: false)
                })
                .Concat(
                    segmentsArray.Skip(index + count)
                );
        }

        /// <summary>
        /// If a single token of type NumericValueToken or StringToken is specified then return a NumericValueExpressionSegment or StringValueExpressionSegment to
        /// represent the constant value (there should be zero arguments in this case). If there are two tokens, a KeyWordToken with content "new" and NameToken
        /// (with no arguments) then this can be represented by a NewInstanceExpressionSegment. Otherwise return a CallExpressionSegment. This will throw an
        /// exception if unable to process the content (including cases where the content would result in a compile time error in VBScript).
        /// </summary>
        private static IExpressionSegment GetCallOrNewOrValueExpressionSegment(
            IEnumerable<IToken> tokens,
            IEnumerable<Expression> arguments,
            IToken directedWithReferenceIfAny,
            bool argumentsAreBracketed,
            bool willBeFirstSegmentInCallExpression,
            Action<string> warningLogger)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

            var tokensList = tokens.ToList();
            if (!tokensList.Any())
                throw new ArgumentException("Empty tokens set specified, invalid");
            if (tokensList.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in tokens set");

            // If the first segment in a call expression and the first token is a "." then directedWithReferenceIfAny must be non-null (meaning this statement is
            // found within a "WITH x" construct) or the statement is invalid. The statement "a(0).Name" is broken down into two segments; "a(0)" and ".Name" -
            // clearly starting with a "." is acceptable (mandatory, in fact!) for ".Name" but would only be in the first segment if found within a WITH.
            if (tokensList[0] is MemberAccessorOrDecimalPointToken)
            {
                if (willBeFirstSegmentInCallExpression)
                {
                    if (directedWithReferenceIfAny == null)
                        throw new ArgumentException("The first token in the first segment of an expression can not be a MemberAccessorOrDecimalPointToken unless the statement is found within a WITH construct");
                    tokensList.Insert(0, directedWithReferenceIfAny);
                }
            }
            else if (!willBeFirstSegmentInCallExpression)
                throw new ArgumentException("All segments in a call expression after the first must start with a MemberAccessorOrDecimalPointToken");

            // Any member access involving a numeric literal will result in a compile time error - eg.
            //   WScript.Echo a.1
            //   WScript.Echo a.1()
            //   WScript.Echo 1.a
            if ((tokensList.Count() > 1) && tokensList.Any(t => t is NumericValueToken))
                throw new ArgumentException("Invalid member access - involving numeric literal (this is VBScript compile time error \"Expected end of statement\")");

            // If there are no arguments and no brackets then there's a chance of representing this as a constant-type expression or as a new instance request
            // - If there are brackets following a number literal then it's a runtime error ("Type mismatch")
            // - If there are brackets following a constant then it's a runtime error ("Type mismatch")
            // - If there are brackets following a class instantiation then it's a compile time error
            // Note: This only covers the zero-argument cases with brackets - runtime errors will be raised if "vbObjectError(F1())" or "vbObjectError(F1())"
            // calls were attempted, but these need to be dealt with by the "CALL" method since the arguments must be evaluated before the error raised.
            if (!arguments.Any())
            {
                if (tokensList.Count == 1)
                {
                    var numericValue = tokensList[0] as NumericValueToken;
                    if (numericValue != null)
                    {
                        if (argumentsAreBracketed)
                        {
                            warningLogger("Numeric literal accessed as a method - this will result in a runtime error (line " + (numericValue.LineIndex + 1) + ")");
                            return new RuntimeErrorExpressionSegment(
                                numericValue.Content + "()",
                                new IToken[] { numericValue, new OpenBrace(numericValue.LineIndex), new CloseBrace(numericValue.LineIndex) },
                                typeof(TypeMismatchException),
                                "'[number: " + numericValue.Content + "]' is called like a function"
                            );
                        }
                        return new NumericValueExpressionSegment(numericValue);
                    }
                    var stringValue = tokensList[0] as StringToken;
                    if (stringValue != null)
                    {
                        if (argumentsAreBracketed)
                        {
                            warningLogger("String literal accessed as a method - this will result in a runtime error (line " + (stringValue.LineIndex + 1) + ")");
                            return new RuntimeErrorExpressionSegment(
                                "\"" + stringValue.Content + "\"()",
                                new IToken[] { stringValue, new OpenBrace(stringValue.LineIndex), new CloseBrace(stringValue.LineIndex) },
                                typeof(TypeMismatchException),
                                "'[string: \"" + stringValue.Content + "\"]' is called like a function"
                            );
                        }
                        return new StringValueExpressionSegment(stringValue);
                    }
					var builtInValue = tokensList[0] as BuiltInValueToken;
                    if (builtInValue != null)
                    {
                        if (argumentsAreBracketed)
                        {
                            warningLogger("Built-in constant accessed as a method - this will result in a runtime error (line " + (builtInValue.LineIndex + 1) + ")");
                            return new RuntimeErrorExpressionSegment(
                                builtInValue.Content + "()",
                                new IToken[] { builtInValue, new OpenBrace(builtInValue.LineIndex), new CloseBrace(builtInValue.LineIndex) },
                                typeof(TypeMismatchException),
                                "'" + builtInValue.Content + "' is called like a function"
                            );
                        }
                        return new BuiltInValueExpressionSegment(builtInValue);
                    }
                }
                else if ((tokensList.Count == 2)
                && (tokensList[0] is KeyWordToken)
                && tokensList[0].Content.Equals("new", StringComparison.InvariantCultureIgnoreCase))
                {
                    var newInstanceName = tokensList[1] as NameToken;
                    if (newInstanceName != null)
                    {
                        if (argumentsAreBracketed)
                        {
                            // In VBScript, this is a compile time error (unlike the runtime errors from brackets following the token types above)
                            throw new Exception("Invalid content - \"Expected end of statement\" (there may not be brackets following the class name when using \"new\")");
                        }
                        return new NewInstanceExpressionSegment(newInstanceName);
                    }
                }
            }

			CallExpressionSegment.ArgumentBracketPresenceOptions? zeroArgumentBracketsPresence;
			if (arguments.Any())
				zeroArgumentBracketsPresence = null;
			else if (argumentsAreBracketed)
				zeroArgumentBracketsPresence = CallExpressionSegment.ArgumentBracketPresenceOptions.Present;
			else
				zeroArgumentBracketsPresence = CallExpressionSegment.ArgumentBracketPresenceOptions.Absent;
            return new CallExpressionSegment(
                tokensList.Where(t => !(t is MemberAccessorOrDecimalPointToken)),
                arguments,
				zeroArgumentBracketsPresence
            );
        }

        /// <summary>
		/// Generate a BracketedExpressionSegment instance - if there is only a single expression segment, where that segment is a bracketed segment, then return
		/// that segment rather than wrapping it again (this is done recursively in case there are multiple layers of over-wrapped bracketed segments). Note: If
		/// it ends up that there's only one expression segment in total then this will be returned, unwrapped.
		/// </summary>
        private static IExpressionSegment WrapExpressionSegments(IEnumerable<IExpressionSegment> segments, bool unwrapSingleBracketedTerm)
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
            if ((segmentsArray.Length == 1) && (unwrapSingleBracketedTerm || (segmentsArray[0] is BracketedExpressionSegment)))
            {
                // If we've ended up with a single segment then we might not have to wrap it up further. This is the case if it is a BracketedExpressionSegment
                // (if this is the case then it brackets multiple segments). It may also be the case it the expression is for a simple expression (eg. "a") -
                // in this case we may or may not bracket it up, depending upon the unwrapSingleBracketedTerm argument value; this will vary depending upon
                // whether the simple expression is one side of an operator (in which case bracketing will never be required and so unwrapSingleBracketedTerm
                // will be true) or if it is a function argument (in which case we can't be sure that removing bracketing will have no effect, it may force
                // an argument to be passed ByVal when it would otherwise be ByRef, for example - in this case unwrapSingleBracketedTerm will be false and
                // the single term WILL be bracketed up).
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

            var callSetItemExpressionSegmentBuffer = new List<CallSetItemExpressionSegment>();
            var expressionSegments = new List<IExpressionSegment>();
            foreach (var segment in segments)
            {
                if (segment == null)
                    throw new ArgumentException("Null reference encountered in segments set");

                var callExpressionSegment = segment as CallSetItemExpressionSegment;
                if (callExpressionSegment != null)
                {
                    callSetItemExpressionSegmentBuffer.Add(callExpressionSegment);
                    continue;
                }

                if (callSetItemExpressionSegmentBuffer.Count > 0)
                {
                    if (callSetItemExpressionSegmentBuffer.Count == 1)
                        expressionSegments.Add(callSetItemExpressionSegmentBuffer[0]);
                    else
                        expressionSegments.Add(new CallSetExpressionSegment(callSetItemExpressionSegmentBuffer));
                    callSetItemExpressionSegmentBuffer.Clear();
                }
                expressionSegments.Add(segment);
            }
            if (callSetItemExpressionSegmentBuffer.Count > 0)
            {
                if (callSetItemExpressionSegmentBuffer.Count == 1)
                {
                    // A CallSetItemExpressionSegment should never exist in isolation, so if there is only one segment here it should
                    // be promoted to a CallExpressionSegment (if it wasn't one already)
                    var callExpressionSegment = callSetItemExpressionSegmentBuffer[0];
                    if (!callExpressionSegment.MemberAccessTokens.Any())
                        throw new ArgumentException("Encountered individual CallSetItemExpressionSegment with no Member Access Tokens, zero Member Access Tokens are only allowable with segments are part of a CallSetExpressionSegment (so long as it's not the first segment in that set's content)");
                    expressionSegments.Add(
                        new CallExpressionSegment(
                            callExpressionSegment.MemberAccessTokens,
                            callExpressionSegment.Arguments,
                            callExpressionSegment.ZeroArgumentBracketsPresence
                        )
                    );
                }
                else
                    expressionSegments.Add(new CallSetExpressionSegment(callSetItemExpressionSegmentBuffer));
            }
            return new Expression(expressionSegments);
        }

        /// <summary>
        /// This is used to enforce VBScript's rules of precedence for operators (so that "a + b * c" can be represented as "a + (b * c)"
        /// </summary>
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
