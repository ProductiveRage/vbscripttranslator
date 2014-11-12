using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class TokenBreakerTests
    {
        /// <summary>
        /// Previously, there was an error where a line break would result in a LineIndex increment for both the line break token and the token
        /// preceding it, rather than tokens AFTER the line break
        /// </summary>
        [Fact]
        public void IncrementLineIndexAfterLineBreaks()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("Test", 0),
                    new NameToken("z", 0),
                    new EndOfStatementNewLineToken(0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("Test z\n", 0)),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByWhitespace()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("a", 0),
                    new OperatorToken("&", 0),
                    new NameToken("b", 1)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a & _\nb", 0)),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByTokenBreakers()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("a", 0),
                    new OperatorToken("&", 0),
                    new NameToken("b", 1)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a&_\nb", 0)),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void DoNotConsiderUnderscoresToBeLineContinuationsWhenTheyArePartOfVariableNames()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("a_b", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a_b", 0)),
                new TokenSetComparer()
            );
        }
    }
}
