using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding;
using VBScriptTranslator.StageTwoParser.Tokens;
using VBScriptTranslator.UnitTests.Shared;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.StageTwoParser
{
    public class NumberRebuilderTests
    {
        [Fact]
        public void NegativeOne()
        {
            Assert.Equal(
                new[]
                {
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void BracketedNegativeOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new OpenBrace(0),
                    new NumericValueToken("-1", 0),
                    new CloseBrace(0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OpenBrace(0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void PointOne()
        {
            Assert.Equal(
                new[]
                {
                    new NumericValueToken("0.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void OnePointOne()
        {
            Assert.Equal(
                new[]
                {
                    new NumericValueToken("1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void NegativeOnePointOne()
        {
            Assert.Equal(
                new[]
                {
                    new NumericValueToken("-1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void NegativePointOne()
        {
            Assert.Equal(
                new[]
                {
                    new NumericValueToken("-0.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void OnePlusNegativeOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("+", 0),
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new OperatorToken("+", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void NegativeOneAsNonBracketedArgument()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("fnc", 0),
                    new NumericValueToken("1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NameToken("fnc", 0),
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void PointOneAsNonBracketedArgument()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("fnc", 0),
                    new NumericValueToken("0.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NameToken("fnc", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void ForLoopWithNegativeConstraints()
        {
            Assert.Equal(
                new IToken[]
                {
                    new KeyWordToken("FOR", 0),
                    new NameToken("i", 0),
                    new ComparisonOperatorToken("=", 0),
                    new NumericValueToken("-1", 0),
                    new KeyWordToken("TO", 0),
                    new NumericValueToken("-4", 0),
                    new KeyWordToken("STEP", 0),
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new KeyWordToken("FOR", 0),
                        new NameToken("i", 0),
                        new ComparisonOperatorToken("=", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new KeyWordToken("TO", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("4", 0),
                        new KeyWordToken("STEP", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// When NameTokens are prefixed with a MemberAccessorOrDecimalPointToken, this is presumably because the content is wrapped in a "WITH" statement
        /// that will resolve the property / method access. As such, it shouldn't be assumed that a trailing dot is always a decimal point.
        /// </summary>
        [Fact]
        public void DoNotTryToTreatMemberSeparatorRelyUponWithKeywordAsDecimalPoint()
        {
            Assert.Equal(
                new IToken[]
                {
                    new MemberAccessorToken(0),
                    new NameToken("Name", 0),
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NameToken("Name", 0),
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
