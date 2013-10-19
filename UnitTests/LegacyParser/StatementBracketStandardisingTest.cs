using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.LegacyParser.Helpers;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class StatementBracketStandardisingTest
    {
        public class NonChangingTests
        {
            [Fact]
            public void DirectMethodWithNoArgumentsAndNoBrackets()
            {
                var tokens = new[]
                {
                    BaseAtomTokenGenerator.Get("Test")
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void DirectMethodWithSingleArgumentWithBrackets()
            {
                var tokens = new[]
                {
                    BaseAtomTokenGenerator.Get("Test"),
                    new OpenBrace("("),
                    BaseAtomTokenGenerator.Get("1"),
                    new CloseBrace(")")
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void DirectMethodWithSingleArgumentWithBracketsAndCallKeyword()
            {
                var tokens = new[]
                {
                    BaseAtomTokenGenerator.Get("Test"),
                    new OpenBrace("("),
                    BaseAtomTokenGenerator.Get("1"),
                    new CloseBrace(")")
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Present);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void DirectMethodWithSingleArgumentWithBracketsAndCallKeywordAndExtraBracketsCall()
            {
                // eg. "CALL(Test(1))"
                // Note that "(Test(1))" is not valid but the extra brackets are allowable when CALL is used
                var tokens = new[]
                {
                    new OpenBrace("("),
                    BaseAtomTokenGenerator.Get("Test"),
                    new OpenBrace("("),
                    BaseAtomTokenGenerator.Get("1"),
                    new CloseBrace(")"),
                    new CloseBrace(")")
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Present);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void ObjectMethodWithSingleArgumentWithBrackets()
            {
                var tokens = new[]
                {
                    BaseAtomTokenGenerator.Get("a"),
                    new MemberAccessorOrDecimalPointToken("."),
                    BaseAtomTokenGenerator.Get("Test"),
                    new OpenBrace("("),
                    BaseAtomTokenGenerator.Get("1"),
                    new CloseBrace(")")
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }

        public class ChangingTests
        {
            [Fact]
            public void DirectMethodWithSingleArgumentWithoutBrackets()
            {
                var statement = new Statement(
                    new[]
                    {
                        BaseAtomTokenGenerator.Get("Test"),
                        BaseAtomTokenGenerator.Get("1")
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new[]
                    {
                        BaseAtomTokenGenerator.Get("Test"),
                        new OpenBrace("("),
                        BaseAtomTokenGenerator.Get("1"),
                        new CloseBrace(")")
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void ObjectMethodWithSingleArgumentWithoutBrackets()
            {
                var statement = new Statement(
                    new[]
                    {
                        BaseAtomTokenGenerator.Get("a"),
                        new MemberAccessorOrDecimalPointToken("."),
                        BaseAtomTokenGenerator.Get("Test"),
                        BaseAtomTokenGenerator.Get("1")
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new[]
                    {
                        BaseAtomTokenGenerator.Get("a"),
                        new MemberAccessorOrDecimalPointToken("."),
                        BaseAtomTokenGenerator.Get("Test"),
                        new OpenBrace("("),
                        BaseAtomTokenGenerator.Get("1"),
                        new CloseBrace(")")
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }
    }
}
