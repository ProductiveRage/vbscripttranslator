using System;
using VBScriptTranslator.UnitTests.LegacyParser.Helpers;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
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
                    GetAtomToken("Test")
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
                    GetAtomToken("Test"),
                    new OpenBrace("("),
                    GetAtomToken("1"),
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
                    GetAtomToken("Test"),
                    new OpenBrace("("),
                    GetAtomToken("1"),
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
                    GetAtomToken("Test"),
                    new OpenBrace("("),
                    GetAtomToken("1"),
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
                    GetAtomToken("a"),
                    new MemberAccessorOrDecimalPointToken("."),
                    GetAtomToken("Test"),
                    new OpenBrace("("),
                    GetAtomToken("1"),
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
                        GetAtomToken("Test"),
                        GetAtomToken("1")
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new[]
                    {
                        GetAtomToken("Test"),
                        new OpenBrace("("),
                        GetAtomToken("1"),
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
                        GetAtomToken("a"),
                        new MemberAccessorOrDecimalPointToken("."),
                        GetAtomToken("Test"),
                        GetAtomToken("1")
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new[]
                    {
                        GetAtomToken("a"),
                        new MemberAccessorOrDecimalPointToken("."),
                        GetAtomToken("Test"),
                        new OpenBrace("("),
                        GetAtomToken("1"),
                        new CloseBrace(")")
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }

        private static AtomToken GetAtomToken(string content)
        {
            var token = AtomToken.GetNewToken(content);
            if (token.GetType() != typeof(AtomToken))
                throw new ArgumentException("Specified content was not mapped to an AtomToken, it was mapped to " + token.GetType());
            return (AtomToken)token;
        }
    }
}
