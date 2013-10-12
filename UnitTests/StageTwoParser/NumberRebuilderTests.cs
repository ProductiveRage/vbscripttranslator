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
                    new NumericValueToken(-1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        new OperatorToken("-"),
                        Misc.GetAtomToken("1")
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
                    new OpenBrace("("),
                    new NumericValueToken(-1),
                    new CloseBrace(")")
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        new OpenBrace("("),
                        new OperatorToken("-"),
                        Misc.GetAtomToken("1"),
                        new CloseBrace(")")
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
                    new NumericValueToken(0.1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        new MemberAccessorOrDecimalPointToken("."),
                        Misc.GetAtomToken("1")
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
                    new NumericValueToken(1.1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        Misc.GetAtomToken("1"),
                        new MemberAccessorOrDecimalPointToken("."),
                        Misc.GetAtomToken("1")
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
                    new NumericValueToken(-1.1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        new OperatorToken("-"),
                        Misc.GetAtomToken("1"),
                        new MemberAccessorOrDecimalPointToken("."),
                        Misc.GetAtomToken("1")
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
                    new NumericValueToken(-0.1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        new OperatorToken("-"),
                        new MemberAccessorOrDecimalPointToken("."),
                        Misc.GetAtomToken("1")
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
                    new NumericValueToken(1),
                    new OperatorToken("+"),
                    new NumericValueToken(-1)
                },
                NumberRebuilder.Rebuild(
                    new[]
                    {
                        Misc.GetAtomToken("1"),
                        new OperatorToken("+"),
                        new OperatorToken("-"),
                        Misc.GetAtomToken("1")
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
