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
                    new UnprocessedContentToken("strValue = "),
                    new StringToken("Test string with \"quoted\" content"),
                    new UnprocessedContentToken("\n")
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
                    new EscapedNameToken("[]"),
                    new UnprocessedContentToken(" = 1")
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
                    new UnprocessedContentToken("Dim "),
                    new EscapedNameToken("[]"),
                    new UnprocessedContentToken(": "),
                    new EscapedNameToken("[]"),
                    new UnprocessedContentToken(" = 1")
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
                    new UnprocessedContentToken("WScript.Echo 1"),
                    new EndOfStatementSameLineToken(),
                    new InlineCommentToken(" Test")
                },
                StringBreaker.SegmentString(
                    "WScript.Echo 1 ' Test"
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
                    new CommentToken(" Test"),
                    new UnprocessedContentToken("WScript.Echo 1")
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
                    new UnprocessedContentToken("WScript.Echo 1"),
                    new EndOfStatementSameLineToken(),
                    new InlineCommentToken(" Test")
                },
                StringBreaker.SegmentString(
                    "WScript.Echo 1 REM Test"
                ),
                new TokenSetComparer()
            );
        }
    }
}
