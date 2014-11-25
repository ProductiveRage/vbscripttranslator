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

        /// <summary>
        /// I realised that "1/0" wasn't being correctly broken down since the "/" wasn't being considered a "Token Break Character" and so the "1/0" was being
        /// interpreted as a NameToken, instead of two numeric value tokens and an operator.
        /// </summary>
        [Fact]
        public void EnsureThatDivisionOperatorsAreRecognised()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1, 0),
                    new OperatorToken("/", 0),
                    new NumericValueToken(0, 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1/0", 0)),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This is the same issue as that for which the EnsureThatDivisionOperatorsAreRecognised test was added, but for the integer division opereator (back
        /// slash, rather than forward)
        /// </summary>
        [Fact]
        public void EnsureThatIntegerDivisionOperatorsAreRecognised()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken(1, 0),
                    new OperatorToken("\\", 0),
                    new NumericValueToken(0, 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1\\0", 0)),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This is an issue identified in testing real content. The first of the following was correctly parsed while the second wasn't -
        ///   value <> ""
        ///   value<> ""
        /// It should be broken down into four tokens:
        ///   NameToken:"value"
        ///   ComparisonOperationToken:"<"
        ///   ComparisonOperationToken:">"
        ///   StringToken:""
        /// The TokenBreaker would then get an UnprocessedContentToken with content "value<>" which it needs to break into three.
        /// </summary>
        [Fact]
        public void LessThanComparisonOperatorIndicatesTokenBreakRegardlessOfWhitespace()
        {
            Assert.Equal(
                new IToken[]
                {
                    new NameToken("value", 0),
                    new ComparisonOperatorToken("<", 0),
                    new ComparisonOperatorToken(">", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("value<>", 0)),
                new TokenSetComparer()
            );
        }
    }
}
