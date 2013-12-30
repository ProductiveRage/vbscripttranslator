using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class StringBreakerTests
    {
        [Fact]
        public void VariableSetToStringContentIncludedQuotedContent()
        {
            Assert.Equal(
                new IToken[]
                {
                    new UnprocessedContentToken("strValue = ", 0),
                    new StringToken("Test string with \"quoted\" content", 0),
                    new UnprocessedContentToken("\n", 0)
                },
                StringBreaker.SegmentString(
                    "strValue = \"Test string with \"\"quoted\"\" content\"\n"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This tests the minimum escaped-content variable name that is possible (a blank variable name, escaped by square brackets)
        /// </summary>
        [Fact]
        public void EmptyContentEscapedVariableNameIsSetToNumericValue()
        {
            Assert.Equal(
                new IToken[]
                {
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(" = 1", 0)
                },
                StringBreaker.SegmentString(
                    "[] = 1"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This tests the minimum escaped-content variable name that is possible (a blank variable name, escaped by square brackets)
        /// </summary>
        [Fact]
        public void DeclaredEmptyContentEscapedVariableNameIsSetToNumericValue()
        {
            Assert.Equal(
                new IToken[]
                {
                    new UnprocessedContentToken("Dim ", 0),
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(": ", 0),
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(" = 1", 0)
                },
                StringBreaker.SegmentString(
                    "Dim []: [] = 1"
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void InlineCommentsAreIdentifiedAsSuch()
        {
            // The StringBreaker will insert an EndOfStatementSameLineToken between the UnprocessedContentToken and InlineCommentToken
            // since that the later processes rely on end-of-statement tokens, even before an inline comment
            Assert.Equal(
                new IToken[]
                {
                    new UnprocessedContentToken("WScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 0)
                },
                StringBreaker.SegmentString(
                    "WScript.Echo 1 ' Test"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This recreates a bug where if there were line returns in the unprocessed content before what should be an inline comment, it
        /// wasn't realised that these were before the line that the comment should be inline with
        /// </summary>
        [Fact]
        public void InlineCommentsAreIdentifiedAsSuchWhenAfterMultipleLinesOfContent()
        {
            // The StringBreaker will insert an EndOfStatementSameLineToken between the UnprocessedContentToken and InlineCommentToken
            // since that the later processes rely on end-of-statement tokens, even before an inline comment
            Assert.Equal(
                new IToken[]
                {
                    new UnprocessedContentToken("\nWScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 1)
                },
                StringBreaker.SegmentString(
                    "\nWScript.Echo 1 ' Test"
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void REMCommentsAreIdentified()
        {
            Assert.Equal(
                new IToken[]
                {
                    new CommentToken(" Test", 0),
                    new UnprocessedContentToken("WScript.Echo 1", 1)
                },
                StringBreaker.SegmentString(
                    "REM Test\nWScript.Echo 1"
                ),
                new TokenSetComparer()
            );
        }

        [Fact]
        public void InlineREMCommentsAreIdentified()
        {
            Assert.Equal(
                new IToken[]
                {
                    new UnprocessedContentToken("WScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 0)
                },
                StringBreaker.SegmentString(
                    "WScript.Echo 1 REM Test"
                ),
                new TokenSetComparer()
            );
        }
    }
}
