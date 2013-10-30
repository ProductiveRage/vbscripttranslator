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
                    new NumericValueToken(1, 0),
                    new OperatorToken("-", 0),
                    new NumericValueToken(1, 0)
                },
                OperatorCombiner.Combine(
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

        [Fact]
        public void OneMinusNegativeOne()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1, 0),
                    new OperatorToken("+", 0),
                    new NumericValueToken(1, 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(1, 0),
                        new OperatorToken("-", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(1, 0),
                    new OperatorToken("*", 0),
                    new NumericValueToken(1, 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(1, 0),
                        new OperatorToken("*", 0),
                        new OperatorToken("+", 0),
                        new NumericValueToken(1, 0)
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
                    new NumericValueToken(2, 0),
                    new ComparisonOperatorToken(">=", 0),
                    new NumericValueToken(1, 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken(2, 0),
                        new ComparisonOperatorToken(">", 0),
                        new ComparisonOperatorToken("=", 0),
                        new NumericValueToken(1, 0)
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
