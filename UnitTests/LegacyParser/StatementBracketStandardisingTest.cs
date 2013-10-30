using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
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
                    new NameToken("Test", 0)
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
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NumericValueToken(1, 0),
                    new CloseBrace(0)
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
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NumericValueToken(1, 0),
                    new CloseBrace(0)
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
                var tokens = new IToken[]
                {
                    new OpenBrace(0),
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NumericValueToken(1, 0),
                    new CloseBrace(0),
                    new CloseBrace(0)
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
                var tokens = new IToken[]
                {
                    new NameToken("a", 0),
                    new MemberAccessorOrDecimalPointToken(".", 0),
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NumericValueToken(1, 0),
                    new CloseBrace(0)
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
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new NumericValueToken(1, 0)
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken(1, 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void ObjectMethodWithSingleArgumentWithoutBrackets()
            {
                var statement = new Statement(
                    new IToken[]
                    {
                        new NameToken("a", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NameToken("Test", 0),
                        new NumericValueToken(1, 0)
                    },
                    Statement.CallPrefixOptions.Absent
                );
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("a", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken(1, 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }
    }
}
