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
                        CALL(Misc.GetAtomToken("Test"))
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("Test")
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
                        CALL(Misc.GetAtomToken("Test"))
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("Test"),
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
                            new[] { Misc.GetAtomToken("a"), Misc.GetAtomToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("a"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("Test")
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
                            new[] { Misc.GetAtomToken("a"), Misc.GetAtomToken("b"), Misc.GetAtomToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("a"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("b"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("Test")
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
                            new[] { Misc.GetAtomToken("Test") },
                            new[] { Misc.GetAtomToken("1") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("Test"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("1"),
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
                            new[] { Misc.GetAtomToken("Test") },
                            new[] { Misc.GetAtomToken("1") },
                            new[] { Misc.GetAtomToken("2") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("Test"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("1"),
                        new ArgumentSeparatorToken(","),
                        Misc.GetAtomToken("2"),
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
                            new[] { Misc.GetAtomToken("Test") },
                            EXP(
                                CALL(
                                    new[] { Misc.GetAtomToken("Test2") },
                                    new[] { Misc.GetAtomToken("1") }
                                )
                            ),
                            EXP(CALL(Misc.GetAtomToken("2")))
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("Test"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("Test2"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("1"),
                        new CloseBrace(")"),
                        new ArgumentSeparatorToken(","),
                        Misc.GetAtomToken("2"),
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
                        CALL(
                            new[] { Misc.GetAtomToken("a") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        CALL(
                            new[] { Misc.GetAtomToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("a"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("0"),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("Test")
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
                        CALL(
                            new[] { Misc.GetAtomToken("a"), Misc.GetAtomToken("b") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        CALL(
                            new[] { Misc.GetAtomToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("a"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("b"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("0"),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("Test")
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
                        CALL(
                            new[] { Misc.GetAtomToken("a") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        CALL(
                            new[] { Misc.GetAtomToken("b"), Misc.GetAtomToken("Test") }
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                        Misc.GetAtomToken("a"),
                        new OpenBrace("("),
                        Misc.GetAtomToken("0"),
                        new CloseBrace(")"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("b"),
                        new MemberAccessorToken(),
                        Misc.GetAtomToken("Test")
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
                            EXP(
                                CALL(Misc.GetAtomToken("a")),
                                OP(new OperatorToken("+")),
                                CALL(Misc.GetAtomToken("b"))
                            )
                        ),
                        OP(new OperatorToken("+")),
                        CALL(Misc.GetAtomToken("c"))
                    )
                },
                ExpressionGenerator.Generate(new[] {
                    Misc.GetAtomToken("a"),
                    new OperatorToken("+"),
                    Misc.GetAtomToken("b"),
                    new OperatorToken("+"),
                    Misc.GetAtomToken("c")
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
                        CALL(Misc.GetAtomToken("a")),
                        OP(new OperatorToken("+")),
                        BR(
                            EXP(
                                CALL(Misc.GetAtomToken("b")),
                                OP(new OperatorToken("*")),
                                CALL(Misc.GetAtomToken("c"))
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                    Misc.GetAtomToken("a"),
                    new OperatorToken("+"),
                    Misc.GetAtomToken("b"),
                    new OperatorToken("*"),
                    Misc.GetAtomToken("c")
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
            // a + b * c
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(Misc.GetAtomToken("a")),
                        OP(new ComparisonToken("=")),
                        BR(
                            EXP(
                                CALL(Misc.GetAtomToken("b")),
                                OP(new OperatorToken("+")),
                                CALL(Misc.GetAtomToken("c"))
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new[] {
                    Misc.GetAtomToken("a"),
                    new ComparisonToken("="),
                    Misc.GetAtomToken("b"),
                    new OperatorToken("+"),
                    Misc.GetAtomToken("c")
                }),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(IEnumerable<Expression> expressions)
        {
            return new BracketedExpressionSegment(expressions);
        }

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(params Expression[] expressions)
        {
            return new BracketedExpressionSegment((IEnumerable<Expression>)expressions);
        }

        /// <summary>
        /// Create an CallExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static CallExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments)
        {
            return new CallExpressionSegment(
                memberAccessTokens,
                arguments
            );
        }

        /// <summary>
        /// Create a CallExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static CallExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params Expression[] arguments)
        {
            return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions
        /// </summary>
        private static CallExpressionSegment CALL(IToken memberAccessToken, params Expression[] arguments)
        {
            return CALL(new[] { memberAccessToken }, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions expressed as token sets
        /// </summary>
        private static CallExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
        {
            return CALL(
                memberAccessTokens,
                arguments.Select(a => new Expression(new[] { CALL(a) }))
            );
        }

        private static OperatorOrComparisonExpressionSegment OP(IToken token)
        {
            return new OperatorOrComparisonExpressionSegment(token);
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
