using System;
using VBScriptTranslator.UnitTests.LegacyParser;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;
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
                    GetAtomToken("1"),
                    new OperatorToken("-"),
                    GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        GetAtomToken("1"),
                        new OperatorToken("+"),
                        new OperatorToken("-"),
                        GetAtomToken("1")
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
                    GetAtomToken("1"),
                    new OperatorToken("+"),
                    GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        GetAtomToken("1"),
                        new OperatorToken("-"),
                        new OperatorToken("-"),
                        GetAtomToken("1")
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
                    GetAtomToken("1"),
                    new OperatorToken("*"),
                    GetAtomToken("1")
                },
                OperatorCombiner.Combine(
                    new[]
                    {
                        GetAtomToken("1"),
                        new OperatorToken("*"),
                        new OperatorToken("+"),
                        GetAtomToken("1")
                    }
                ),
                new TokenSetComparer()
            );
        }

        private AtomToken GetAtomToken(string content)
        {
            var token = AtomToken.GetNewToken(content);
            if (token.GetType() != typeof(AtomToken))
                throw new ArgumentException("Specified content was not mapped to an AtomToken, it was mapped to " + token.GetType());
            return (AtomToken)token;
        }
    }
}
