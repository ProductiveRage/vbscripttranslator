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
                    new NumericValueToken(-1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(-1, 0),
                    new CloseBrace(0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OpenBrace(0),
                        new OperatorToken("-", 0),
                        new NumericValueToken(1, 0),
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
                    new NumericValueToken(0.1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(1.1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NumericValueToken(1, 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(-1.1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken(1, 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(-0.1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(1, 0),
                    new OperatorToken("+", 0),
                    new NumericValueToken(-1, 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NumericValueToken(1, 0),
                        new OperatorToken("+", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken(1, 0)
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
