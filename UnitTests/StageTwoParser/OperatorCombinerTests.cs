using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;
using VBScriptTranslator.UnitTests.Shared;
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
                new[]
                {
                    Misc.GetAtomToken("1"),
                    new OperatorToken("-"),
                    Misc.GetAtomToken("1")
                },
                OperatorCombiner.Combine(
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

        [Fact]
        public void OneMinusNegativeOne()
        {
            Assert.Equal(
                new[]
                {
                    Misc.GetAtomToken("1"),
                    new OperatorToken("+"),
                    Misc.GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        Misc.GetAtomToken("1"),
                        new OperatorToken("-"),
                        new OperatorToken("-"),
                        Misc.GetAtomToken("1")
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void OneMultipliedByPlusOne()
        {
            Assert.Equal(
                new[]
                {
                    Misc.GetAtomToken("1"),
                    new OperatorToken("*"),
                    Misc.GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        Misc.GetAtomToken("1"),
                        new OperatorToken("*"),
                        new OperatorToken("+"),
                        Misc.GetAtomToken("1")
                    }
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void TwoGreaterThanOrEqualToOne()
        {
            Assert.Equal(
                new[]
                {
                    Misc.GetAtomToken("2"),
                    new ComparisonOperatorToken(">="),
                    Misc.GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        Misc.GetAtomToken("2"),
                        new ComparisonOperatorToken(">"),
                        new ComparisonOperatorToken("="),
                        Misc.GetAtomToken("1")
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
