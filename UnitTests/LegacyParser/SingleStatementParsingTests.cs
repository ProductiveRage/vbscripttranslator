using System;
using VBScriptTranslator.UnitTests.LegacyParser.Helpers;
using VBScriptTranslator.LegacyParser;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
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
                        new[]
                        {
                            GetAtomToken("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            GetAtomToken("Echo"),
                            GetAtomToken("1"),
                            new MemberAccessorOrDecimalPointToken("."),
                            GetAtomToken("1")
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
                        new[]
                        {
                            GetAtomToken("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            GetAtomToken("Echo"),
                            new OpenBrace("("),
                            GetAtomToken("1"),
                            new MemberAccessorOrDecimalPointToken("."),
                            GetAtomToken("1"),
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
                            GetAtomToken("a"),
                        },
                        new[]
                        {
                            GetAtomToken("Nothing"),
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
                        new[]
                        {
                            GetAtomToken("a"),
                            new OpenBrace("("),
                            GetAtomToken("0"),
                            new ArgumentSeparatorToken(","),
                            GetAtomToken("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            GetAtomToken("1"),
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
                        new[]
                        {
                            GetAtomToken("a"),
                            new OpenBrace("("),
                            GetAtomToken("0"),
                            new ArgumentSeparatorToken(","),
                            GetAtomToken("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            GetAtomToken("1"),
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
                        new[]
                        {
                            GetAtomToken("a"),
                            new OpenBrace("("),
                            GetAtomToken("0"),
                            new ArgumentSeparatorToken(","),
                            GetAtomToken("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            GetAtomToken("Nothing"),
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
                        new[]
                        {
                            GetAtomToken("a"),
                            new OpenBrace("("),
                            GetAtomToken("GetValue"),
                            new OpenBrace("("),
                            GetAtomToken("1"),
                            new ArgumentSeparatorToken(","),
                            GetAtomToken("3"),
                            new CloseBrace(")"),
                            new ArgumentSeparatorToken(","),
                            GetAtomToken("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            GetAtomToken("Nothing"),
                        },
                        ValueSettingStatement.ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse("Set a(GetValue(1, 3), 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        private AtomToken GetAtomToken(string content)
        {
            var token = AtomToken.GetNewToken(content);
            if (token.GetType() != typeof(AtomToken))
                throw new ArgumentException("Specified content was not mapped to an AtomToken, it was mapped to " + token.GetType());
            return (AtomToken)token;
        }
    }
}
