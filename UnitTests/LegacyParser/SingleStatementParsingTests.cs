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
                            BaseAtomTokenGenerator.Get("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            BaseAtomTokenGenerator.Get("Echo"),
                            BaseAtomTokenGenerator.Get("1"),
                            new MemberAccessorOrDecimalPointToken("."),
                            BaseAtomTokenGenerator.Get("1")
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
                            BaseAtomTokenGenerator.Get("WScript"),
                            new MemberAccessorOrDecimalPointToken("."),
                            BaseAtomTokenGenerator.Get("Echo"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("1"),
                            new MemberAccessorOrDecimalPointToken("."),
                            BaseAtomTokenGenerator.Get("1"),
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
                            BaseAtomTokenGenerator.Get("a"),
                        },
                        new[]
                        {
                            BaseAtomTokenGenerator.Get("Nothing"),
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
                            BaseAtomTokenGenerator.Get("a"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("0"),
                            new ArgumentSeparatorToken(","),
                            BaseAtomTokenGenerator.Get("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            BaseAtomTokenGenerator.Get("1"),
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
                            BaseAtomTokenGenerator.Get("a"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("0"),
                            new ArgumentSeparatorToken(","),
                            BaseAtomTokenGenerator.Get("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            BaseAtomTokenGenerator.Get("1"),
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
                            BaseAtomTokenGenerator.Get("a"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("0"),
                            new ArgumentSeparatorToken(","),
                            BaseAtomTokenGenerator.Get("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            BaseAtomTokenGenerator.Get("Nothing"),
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
                            BaseAtomTokenGenerator.Get("a"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("GetValue"),
                            new OpenBrace("("),
                            BaseAtomTokenGenerator.Get("1"),
                            new ArgumentSeparatorToken(","),
                            BaseAtomTokenGenerator.Get("3"),
                            new CloseBrace(")"),
                            new ArgumentSeparatorToken(","),
                            BaseAtomTokenGenerator.Get("0"),
                            new CloseBrace(")"),
                        },
                        new[]
                        {
                            BaseAtomTokenGenerator.Get("Nothing"),
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
