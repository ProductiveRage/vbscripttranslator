using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser;
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
                    E(
                        S(Misc.GetAtomToken("Test"))
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
                    E(
                        S(Misc.GetAtomToken("Test"))
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
                    E(
                        S(
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
                    E(
                        S(
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
                    E(
                        S(
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
        public void DirectFunctionCallWithTwoArgumentsOneIsNestedDirectionFunctionCallWithOneArgument()
        {
            // Test(Test2(1), 2)
            Assert.Equal(new[]
                {
                    E(
                        S(
                            new[] { Misc.GetAtomToken("Test") },
                            E(
                                S(
                                    new[] { Misc.GetAtomToken("Test2") },
                                    new[] { Misc.GetAtomToken("1") }
                                )
                            ),
                            E(S(Misc.GetAtomToken("2")))
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
                    E(
                        S(
                            new[] { Misc.GetAtomToken("a") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        S(
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
                    E(
                        S(
                            new[] { Misc.GetAtomToken("a"), Misc.GetAtomToken("b") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        S(
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
                    E(
                        S(
                            new[] { Misc.GetAtomToken("a") },
                            new[] { Misc.GetAtomToken("0") }
                        ),
                        S(
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
        /// Create an ExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static ExpressionSegment S(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments)
        {
            return new ExpressionSegment(
                memberAccessTokens,
                arguments
            );
        }

        /// <summary>
        /// Create an ExpressionSegment from member access tokens and argument expressions
        /// </summary>
        private static ExpressionSegment S(IEnumerable<IToken> memberAccessTokens, params Expression[] arguments)
        {
            return S(memberAccessTokens, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create an ExpressionSegment from a single member access token and argument expressions
        /// </summary>
        private static ExpressionSegment S(IToken memberAccessToken, params Expression[] arguments)
        {
            return S(new[] { memberAccessToken }, (IEnumerable<Expression>)arguments);
        }

        /// <summary>
        /// Create an ExpressionSegment from a single member access token and argument expressions expressed as token sets
        /// </summary>
        private static ExpressionSegment S(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
        {
            return S(
                memberAccessTokens,
                arguments.Select(a => new Expression(new[] { S(a) }))
            );
        }

        /// <summary>
        /// Create an Expression from multiple ExpressionSegments
        /// </summary>
        private static Expression E(params ExpressionSegment[] segments)
        {
            return new Expression(segments);
        }

        /// <summary>
        /// Create an Expression from a single ExpressionSegment
        /// </summary>
        private static Expression E(ExpressionSegment segment)
        {
            return E(new[] { segment });
        }
    }
}
