using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.StageTwoParser
{
    public class OperatorCombinerTests
    {
        [Fact]
        public void OnePlusNegativeOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1),
                    new OperatorToken("-"),
                    new NumericValueToken(1)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(1),
                        new OperatorToken("+"),
                        new OperatorToken("-"),
                        new NumericValueToken(1)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void OneMinusNegativeOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1),
                    new OperatorToken("+"),
                    new NumericValueToken(1)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(1),
                        new OperatorToken("-"),
                        new OperatorToken("-"),
                        new NumericValueToken(1)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void OneMultipliedByPlusOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1),
                    new OperatorToken("*"),
                    new NumericValueToken(1)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(1),
                        new OperatorToken("*"),
                        new OperatorToken("+"),
                        new NumericValueToken(1)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void TwoGreaterThanOrEqualToOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(2),
                    new ComparisonOperatorToken(">="),
                    new NumericValueToken(1)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(2),
                        new ComparisonOperatorToken(">"),
                        new ComparisonOperatorToken("="),
                        new NumericValueToken(1)
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
