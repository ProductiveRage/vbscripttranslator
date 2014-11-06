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
            public void FunctionCallWithCallKeywordAndMandatoryArgumentBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NameToken("a", 0),
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
            public void FunctionCallWithCallKeywordAndMandatoryAndAdditionalArgumentBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new OpenBrace(0),
                    new NameToken("a", 0),
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
            public void FunctionCallWithCallKeywordAndMandatoryArgumentBracketsAroundNegatedVariable()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new OperatorToken("-", 0),
                    new NameToken("a", 0),
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
            public void FunctionCallWithCallKeywordAndMandatoryArgumentBracketsAroundNegatedConstant()
            {
                // For statements such as "CALL Test(-1)" the "-" and "1" tokens should have been combined before passing them to the
                // Statement constructor. This is unlike "Test -1" since the logic for ensuring that this is considered "Test(-1)"
                // rather than the value "Test" minus one is in the Statement class, work before that point is unaware of it.
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NumericValueToken(-1, 0),
                    new CloseBrace(0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Present);
                Assert.Equal(
                    tokens,
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }

        public class ChangingTests
        {
            /// <summary>
            /// The statement "Test(a)" treats the brackets as passing "a" ByVal to Test, even if the argument for Test is specified as ByRef,
            /// the brackets are not being used to associate the argument(s) with the function call. So additional brackets are required so
            /// that ByVal-enforcing brackets are still present within the standardised format (where brackets MUST be used to associate
            /// arguments with the target function - this is the whole point of the bracket standardising process).
            /// </summary>
            [Fact]
            public void DirectMethodWithSingleArgumentWithBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NameToken("a", 0),
                    new CloseBrace(0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new OpenBrace(0),
                        new NameToken("a", 0),
                        new CloseBrace(0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            /// <summary>
            /// The same logic that applies to DirectMethodWithSingleArgumentWithBrackets applies if an argument already appears to be
            /// "double wrapped", there is no logic to remove extra brackets that are not required
            /// </summary>
            [Fact]
            public void DirectMethodWithSingleArgumentWithDoubleBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new OpenBrace(0),
                    new NameToken("a", 0),
                    new CloseBrace(0),
                    new CloseBrace(0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new OpenBrace(0),
                        new OpenBrace(0),
                        new NameToken("a", 0),
                        new CloseBrace(0),
                        new CloseBrace(0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void DirectMethodWithSingleNegativeArgumentWithoutBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OperatorToken("-", 0),
                    new NameToken("a", 0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new OperatorToken("-", 0),
                        new NameToken("a", 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            /// <summary>
            /// Since "Test -1" is confirmed to mean "Test(-1)" by the Statement bracket-standardising process (rather than possibly
            /// meaning "Test(-1)" or possibly meaning the value "Test" minus one - until the Statement class comes into play, nothing
            /// in the processing chain makes that decision) the "-" and "1" tokens can be combined since it is clear that they are
            /// part of the same value and the "-" token is not an operator acting on "Test" and "1".
            /// </summary>
            [Fact]
            public void DirectMethodWithSingleNegativeNumericArgumentWithoutBrackets()
            {
                // From a previous comment on test "TestMinusOneIsMethodCallNotSubtractionOperation" (which was written but not
                // committed in - and it was a duplicate of this test):
                //   The statement "Test -1" must be interpreted at "Test(-1)", not a subtraction of 1 from Test. If "Test" is an
                //   argument-less function then it will be executed to perform the work of the statement. If it is not (if "Test"
                //   is a variable rather than a function or property) then a runtime error will be raised ("Type mismatch").
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OperatorToken("-", 0),
                    new NumericValueToken(1, 0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken(-1, 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

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
                // This should become "a.Test(1)" in the "standardised" format, the brackets are there for parsing, not for signficant meaning
                // (like forcing the argument to be passed ByVal - see ObjectMethodWithSingleVariableArgumentWithBrackets for that case)
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

            [Fact]
            public void ObjectMethodWithSingleVariableArgumentWithBrackets()
            {
                // This should become "a.Test((b))" since "b" should keep its brackets which indicate the argument be passed ByVal but additional
                // brackets are inserted to "standardise" the format for parsing
                var tokens = new IToken[]
                {
                    new NameToken("a", 0),
                    new MemberAccessorOrDecimalPointToken(".", 0),
                    new NameToken("Test", 0),
                    new OpenBrace(0),
                    new NameToken("b", 0),
                    new CloseBrace(0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("a", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new OpenBrace(0),
                        new NameToken("b", 0),
                        new CloseBrace(0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void DirectMethodWithTwoArgumentsWithoutBrackets()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new NumericValueToken(1, 0),
                    new ArgumentSeparatorToken(",", 0),
                    new NumericValueToken(2, 0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken(1, 0),
                        new ArgumentSeparatorToken(",", 0),
                        new NumericValueToken(2, 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            /// <summary>
            /// As with DirectMethodWithSingleNegativeNumericArgumentWithoutBrackets, a negative sign operator can be combined with a
            /// number to make a negative value if it is the first argument (until the bracket-standardising process, it could have
            /// meant a subtraction operation as far as we knew)
            /// </summary>
            [Fact]
            public void DirectMethodWithTwoArgumentsWithoutBracketsWhereFirstTermIsNegative()
            {
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new OperatorToken("-", 0),
                    new NumericValueToken(1, 0),
                    new ArgumentSeparatorToken(",", 0),
                    new NumericValueToken(2, 0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken(-1, 0),
                        new ArgumentSeparatorToken(",", 0),
                        new NumericValueToken(2, 0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }

            [Fact]
            public void NestedMethodCallWithArgumentsShouldNotGetDoubleBrackets()
            {
                // "Test Test2(a)" should be transformed into "Test(Test2(a))", the brackets around the "a" do not imply that it
                // should be passed ByVal, they are required as the call to Test2 is one where the value is considered (brackets
                // are only only in statements where the return value is not considered - the way in which the value "Test2(a)"
                // is passed to "Test", for example.
                var tokens = new IToken[]
                {
                    new NameToken("Test", 0),
                    new NameToken("Test2", 0),
                    new OpenBrace(0),
                    new NameToken("a", 0),
                    new CloseBrace(0)
                };
                var statement = new Statement(tokens, Statement.CallPrefixOptions.Absent);
                Assert.Equal(
                    new IToken[]
                    {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NameToken("Test2", 0),
                        new OpenBrace(0),
                        new NameToken("a", 0),
                        new CloseBrace(0),
                        new CloseBrace(0)
                    },
                    statement.BracketStandardisedTokens,
                    new TokenSetComparer()
                );
            }
        }
        
        // TODO
        //"Test Test2(1)" // Works - CHANGING
    }
}
