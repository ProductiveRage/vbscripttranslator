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
                            new NameToken("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            new NameToken("Echo"),
                            new NumericValueToken(1),
                            new MemberAccessorOrDecimalPointToken("."),
                            new NumericValueToken(1)
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
                            new NameToken("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            new NameToken("Echo"),
                            new OpenBrace("("),
                            new NumericValueToken(1),
                            new MemberAccessorOrDecimalPointToken("."),
                            new NumericValueToken(1),
                            new CloseBrace(")")
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
                        new[]
                        {
                            new NameToken("a"),
                        },
                        new[]
                        {
                            new BuiltInValueToken("Nothing"),
                        },
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
                        new IToken[]
                        {
                            new NameToken("a"),
                            new OpenBrace("("),
                            new NumericValueToken(0),
                            new ArgumentSeparatorToken(","),
                            new NumericValueToken(0),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            new NumericValueToken(1),
                        },
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
                        new IToken[]
                        {
                            new NameToken("a"),
                            new OpenBrace("("),
                            new NumericValueToken(0),
                            new ArgumentSeparatorToken(","),
                            new NumericValueToken(0),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            new NumericValueToken(1),
                        },
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
                        new IToken[]
                        {
                            new NameToken("a"),
                            new OpenBrace("("),
                            new NumericValueToken(0),
                            new ArgumentSeparatorToken(","),
                            new NumericValueToken(0),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            new BuiltInValueToken("Nothing"),
                        },
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
                        new IToken[]
                        {
                            new NameToken("a"),
                            new OpenBrace("("),
                            new NameToken("GetValue"),
                            new OpenBrace("("),
                            new NumericValueToken(1),
                            new ArgumentSeparatorToken(","),
                            new NumericValueToken(3),
                            new CloseBrace(")"),
                            new ArgumentSeparatorToken(","),
                            new NumericValueToken(0),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            new BuiltInValueToken("Nothing"),
                        },
                        ValueSettingStatement.ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse("Set a(GetValue(1, 3), 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }
    }
}
