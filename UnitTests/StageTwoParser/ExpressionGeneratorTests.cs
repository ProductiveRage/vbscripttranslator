using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using VBScriptTranslator.StageTwoParser.Tokens;
using VBScriptTranslator.UnitTests.Shared;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.StageTwoParser
{
    public class ExpressionGeneratorTests
    {
        [Fact]
        public void DirectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("Test"))
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithNoArgumentsWithBrackets()
        {
            // Test()
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("Test"))
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("Test"),
                        new OpenBrace("("),
                        new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ObjectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // a.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("a"), new NameToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a"),
                        new MemberAccessorToken(),
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void NestedObjectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // a.b.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("a"), new NameToken("b"), new NameToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a"),
                        new MemberAccessorToken(),
                        new NameToken("b"),
                        new MemberAccessorToken(),
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithOneArgument()
        {
            // Test(1)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test") },
                            new[] { new NumericValueToken(1) }
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("Test"),
                        new OpenBrace("("),
                        new NumericValueToken(1),
                        new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithTwoArguments()
        {
            // Test(1, 2)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test") },
                            new[] { new NumericValueToken(1) },
                            new[] { new NumericValueToken(2) }
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("Test"),
                        new OpenBrace("("),
                        new NumericValueToken(1),
                        new ArgumentSeparatorToken(","),
                        new NumericValueToken(2),
                        new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithTwoArgumentsOneIsNestedDirectionFunctionCallWithOneArgument()
        {
            // Test(Test2(1), 2)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test") },
                            EXP(
                                CALL(
                                    new[] { new NameToken("Test2") },
                                    new[] { new NumericValueToken(1) }
                                )
                            ),
                            EXP(CALL(new NumericValueToken(2)))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("Test"),
                        new OpenBrace("("),
                        new NameToken("Test2"),
                        new OpenBrace("("),
                        new NumericValueToken(1),
                        new CloseBrace(")"),
                        new ArgumentSeparatorToken(","),
                        new NumericValueToken(2),
                        new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ArrayElementFunctionCallWithNoArguments()
        {
            // a(0).Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a") },
                                new[] { new NumericValueToken(0) }
                            ),
                            CALL(
                                new[] { new NameToken("Test") }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a"),
                        new OpenBrace("("),
                        new NumericValueToken(0),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ObjectPropertyArrayElementFunctionCallWithNoArguments()
        {
            // a.b(0).Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a"), new NameToken("b") },
                                new[] { new NumericValueToken(0) }
                            ),
                            CALL(
                                new[] { new NameToken("Test") }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a"),
                        new MemberAccessorToken(),
                        new NameToken("b"),
                        new OpenBrace("("),
                        new NumericValueToken(0),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ArrayElementNestedFunctionCallWithNoArguments()
        {
            // a(0).b.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a") },
                                new[] { new NumericValueToken(0) }
                            ),
                            CALL(
                                new[] { new NameToken("b"), new NameToken("Test") }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a"),
                        new OpenBrace("("),
                        new NumericValueToken(0),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        new NameToken("b"),
                        new MemberAccessorToken(),
                        new NameToken("Test")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Additional brackets will be applied around all operations to ensure that VBScript operator rules are always maintained (if the operators
        /// are all equivalent in terms of priority, terms will be bracketed from left-to-right, so a and b should be bracketed together)
        /// </summary>
        [Fact]
        public void AdditionWithThreeTerms()
        {
            // a + b + c
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a")),
                            OP(new OperatorToken("+")),
                            CALL(new NameToken("b"))
                        ),
                        OP(new OperatorToken("+")),
                        CALL(new NameToken("c"))
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new NameToken("b"),
                    new OperatorToken("+"),
                    new NameToken("c")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Multiplication should take precedence over addition so b and c should be bracketed together
        /// </summary>
        [Fact]
        public void AdditionAndMultiplicationWithThreeTerms()
        {
            // a + b * c
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new OperatorToken("+")),
                        BR(
                            CALL(new NameToken("b")),
                            OP(new OperatorToken("*")),
                            CALL(new NameToken("c"))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void AdditionAndMultiplicationWithThreeTermsWhereTheThirdTermIsAnArrayElement()
        {
            // a + b * c(0)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new OperatorToken("+")),
                        BR(
                            CALL(new NameToken("b")),
                            OP(new OperatorToken("*")),
                            CALL(
                                new[] { new NameToken("c") },
                                new[] { new NumericValueToken(0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c"),
                    new OpenBrace("("),
                    new NumericValueToken(0),
                    new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This will try to ensure that the bracket around the array access doesn't interfere with the formatting of the fourth term
        /// </summary>
        [Fact]
        public void AdditionAndMultiplicationAndAdditionWithFourTermsWhereTheThirdTermIsAnArrayElement()
        {
            // a + b * c(0) + d
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a")),
                            OP(new OperatorToken("+")),
                            BR(
                                CALL(new NameToken("b")),
                                OP(new OperatorToken("*")),
                                CALL(
                                    new[] { new NameToken("c") },
                                    new[] { new NumericValueToken(0) }
                                )
                            )
                        ),
                        OP(new OperatorToken("+")),
                        CALL(new NameToken("d"))
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c"),
                    new OpenBrace("("),
                    new NumericValueToken(0),
                    new CloseBrace(")"),
                    new OperatorToken("+"),
                    new NameToken("d"),
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// If an operation is already bracketed then additional brackets should not be added around the operation, they would be unnecessary
        /// </summary>
        [Fact]
        public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketing()
        {
            // a + (b * c)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new OperatorToken("+")),
                        BR(
                            CALL(new NameToken("b")),
                            OP(new OperatorToken("*")),
                            CALL(new NameToken("c"))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new OpenBrace("("),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c"),
                    new CloseBrace(")")
                }),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketingIfTheyAppearInTheMiddleOfTheExpression()
        {
            // a + (b * c) + d
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a")),
                            OP(new OperatorToken("+")),
                            BR(
                                CALL(new NameToken("b")),
                                OP(new OperatorToken("*")),
                                CALL(new NameToken("c"))
                            )
                        ),
                        OP(new OperatorToken("+")),
                        CALL(new NameToken("d"))
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new OpenBrace("("),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c"),
                    new CloseBrace(")"),
                    new OperatorToken("+"),
                    new NameToken("d"),
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Arithmetic operations should take precedence over comparisons so b and c should be bracketed together
        /// </summary>
        [Fact]
        public void AdditionAndEqualityComparisonWithThreeTerms()
        {
            // a = b + c
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new ComparisonOperatorToken("=")),
                        BR(
                            CALL(new NameToken("b")),
                            OP(new OperatorToken("+")),
                            CALL(new NameToken("c"))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new ComparisonOperatorToken("="),
                    new NameToken("b"),
                    new OperatorToken("+"),
                    new NameToken("c")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This covers an array of different types of expression
        /// </summary>
        [Fact]
        public void TestArrayAccessObjectAccessMethodArgumentsMixedArithmeticAndComparisonOperations()
        {
            // a + b * c.d(Test(0), 1) + e = f
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            BR(
                                CALL(new NameToken("a")),
                                OP(new OperatorToken("+")),
                                BR(
                                    CALL(new NameToken("b")),
                                    OP(new OperatorToken("*")),
                                    CALL(
                                        new[] { new NameToken("c"), new NameToken("d") },
                                        EXP(
                                            CALL(
                                                new[] { new NameToken("Test") },
                                                new[] { new NumericValueToken(0) }
                                            )
                                        ),
                                        EXP(CALL(new NumericValueToken(1)))
                                    )
                                )
                            ),
                            OP(new OperatorToken("+")),
                            CALL(new NameToken("e"))
                        ),
                        OP(new ComparisonOperatorToken("=")),
                        CALL(new NameToken("f"))
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("+"),
                    new NameToken("b"),
                    new OperatorToken("*"),
                    new NameToken("c"),
                    new MemberAccessorToken(),
                    new NameToken("d"),
                    new OpenBrace("("),
                    new NameToken("Test"),
                    new OpenBrace("("),
                    new NumericValueToken(0),
                    new CloseBrace(")"),
                    new ArgumentSeparatorToken(","),
                    new NumericValueToken(1),
                    new CloseBrace(")"),
                    new OperatorToken("+"),
                    new NameToken("e"),
                    new ComparisonOperatorToken("="),
                    new NameToken("f")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// To make it clear that the "-" is a one-sided operation (a negation, not a subtraction), it should be bracketed
        /// </summary>
        [Fact]
        public void NegatedTermsShouldBeBracketed()
        {
            // a * -b
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new OperatorToken("*")),
                        BR(
                            OP(new OperatorToken("-")),
                            CALL(new NameToken("b"))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new OperatorToken("*"),
                    new OperatorToken("-"),
                    new NameToken("b")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This is the boolean equivalent of NegatedTermsShouldBeBracketed
        /// </summary>
        [Fact]
        public void LogicalInversionsTermsShouldBeBracketed()
        {
            // a AND NOT b
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a")),
                        OP(new LogicalOperatorToken("AND")),
                        BR(
                            OP(new LogicalOperatorToken("NOT")),
                            CALL(new NameToken("b"))
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                    new NameToken("a"),
                    new LogicalOperatorToken("AND"),
                    new LogicalOperatorToken("NOT"),
                    new NameToken("b")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(IEnumerable<IExpressionSegment> segments)
        {
			return new BracketedExpressionSegment(segments);
        }

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
		private static BracketedExpressionSegment BR(params IExpressionSegment[] segments)
        {
			return new BracketedExpressionSegment((IEnumerable<IExpressionSegment>)segments);
        }

        private static CallSetExpressionSegment CALLSET(params IExpressionSegment[] segments)
        {
            return new CallSetExpressionSegment(segments.Cast<CallExpressionSegment>());
        }

        /// <summary>
        /// Create an CallExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments)
        {
            if ((memberAccessTokens.Count() == 1) && !arguments.Any())
            {
                if (memberAccessTokens.Single() is NumericValueToken)
                    return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
                if (memberAccessTokens.Single() is StringToken)
                    return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
            }
            return new CallExpressionSegment(
                memberAccessTokens,
                arguments
            );
        }

        /// <summary>
        /// Create a CallExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params Expression[] arguments)
        {
            return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions
        /// </summary>
        private static IExpressionSegment CALL(IToken memberAccessToken, params Expression[] arguments)
        {
            return CALL(new[] { memberAccessToken }, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions expressed as token sets
        /// </summary>
        private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
        {
            if ((memberAccessTokens.Count() == 1) && !arguments.Any())
            {
                if (memberAccessTokens.Single() is NumericValueToken)
                    return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
                if (memberAccessTokens.Single() is StringToken)
                    return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
            }
            return CALL(
                memberAccessTokens,
                arguments.Select(a => new Expression(new[] { CALL(a) }))
            );
        }

        private static OperationExpressionSegment OP(OperatorToken token)
        {
            return new OperationExpressionSegment(token);
        }

        /// <summary>
        /// Create an Expression from multiple ExpressionSegments
        /// </summary>
        private static Expression EXP(params IExpressionSegment[] segments)
        {
            return new Expression(segments);
        }

        /// <summary>
        /// Create an Expression from a single ExpressionSegment
        /// </summary>
        private static Expression EXP(IExpressionSegment segment)
        {
            return EXP(new[] { segment });
        }
    }
}
