using VBScriptTranslator.LegacyParser;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.LegacyParser.Helpers;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class SingleStatementParsingTests
    {
        [Fact]
        public void StatementWithMemberAccessAndDecimalValueAndUnwrappedMethodArgument()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new Statement(
                        new IToken[]
                        {
                            new NameToken("WScript", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NameToken("Echo", 0),
                            new NumericValueToken("1", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NumericValueToken("1", 0)
                        },
                        Statement.CallPrefixOptions.Absent
                    )
                },
                Parser.Parse("WScript.Echo 1.1"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void StatementWithMemberAccessAndDecimalValueAndWrappedMethodArgument()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new Statement(
                        new IToken[]
                        {
                            new NameToken("WScript", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NameToken("Echo", 0),
                            new OpenBrace(0),
                            new NumericValueToken("1", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NumericValueToken("1", 0),
                            new CloseBrace(0)
                        },
                        Statement.CallPrefixOptions.Absent
                    )
                },
                Parser.Parse("WScript.Echo(1.1)"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void SingleValueSetToNothing()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new[]
                        {
                            new NameToken("a", 0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSettingStatement.ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse("Set a = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void TwoDimensionalArrayElementSetToNumber()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new NumericValueToken("1", 0),
                        }),
                        ValueSettingStatement.ValueSetTypeOptions.Let
                    )
                },
                Parser.Parse("a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void TwoDimensionalArrayElementSetToNumberWithExplicitLet()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new NumericValueToken("1", 0),
                        }),
                        ValueSettingStatement.ValueSetTypeOptions.Let
                    )
                },
                Parser.Parse("Let a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void TwoDimensionalArrayElementSetToNothing()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSettingStatement.ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse("Set a(0, 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [Fact]
        public void TwoDimensionalArrayElementWithMethodCallIndexSetToNothing()
        {
            Assert.Equal(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NameToken("GetValue", 0),
                            new OpenBrace(0),
                            new NumericValueToken("1", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("3", 0),
                            new CloseBrace(0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSettingStatement.ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse("Set a(GetValue(1, 3), 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }
    }
}
