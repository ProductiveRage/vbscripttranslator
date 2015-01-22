using CSharpSupport.Exceptions;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class NUM
        {
            [Fact]
            public void Empty()
            {
                Assert.Equal(
                    0,
                    GetDefaultRuntimeFunctionalityProvider().NUM(null)
                );
            }

            [Fact]
            public void Null()
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().NUM(DBNull.Value);
                });
            }

            [Fact]
            public void True()
            {
                Assert.Equal(
                    -1,
                    GetDefaultRuntimeFunctionalityProvider().NUM(true)
                );
            }

            [Fact]
            public void False()
            {
                Assert.Equal(
                    0,
                    GetDefaultRuntimeFunctionalityProvider().NUM(false)
                );
            }

            [Fact]
            public void PositiveIntegerString()
            {
                Assert.Equal(
                    12,
                    GetDefaultRuntimeFunctionalityProvider().NUM("12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithLeadingWhitespace()
            {
                Assert.Equal(
                    12,
                    GetDefaultRuntimeFunctionalityProvider().NUM(" 12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithTrailingWhitespace()
            {
                Assert.Equal(
                    12,
                    GetDefaultRuntimeFunctionalityProvider().NUM("12 ")
                );
            }

            [Fact]
            public void TwoIntegersSeparatedByWhitespace()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().NUM("1 1");
                });
            }

            [Fact]
            public void PositiveDecimalString()
            {
                Assert.Equal(
                    1.2,
                    GetDefaultRuntimeFunctionalityProvider().NUM("1.2")
                );
            }

            [Fact]
            public void PseudoNumberWithMultipleDecimalPoints()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().NUM("1.1.0");
                });
            }

            // TODO: String with underscores, hashes, exclamation marks

            // TODO: Negatives
            // TODO: Multiple negatives
        }
    }
}
