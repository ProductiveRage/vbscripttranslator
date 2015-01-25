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
                    (Int16)0,
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
                    (Int16)(-1),
                    GetDefaultRuntimeFunctionalityProvider().NUM(true)
                );
            }

            [Fact]
            public void False()
            {
                Assert.Equal(
                    (Int16)0,
                    GetDefaultRuntimeFunctionalityProvider().NUM(false)
                );
            }

            [Fact]
            public void BlankString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().NUM("");
                });
            }

            [Fact]
            public void PositiveIntegerString()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    GetDefaultRuntimeFunctionalityProvider().NUM("12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithLeadingWhitespace()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    GetDefaultRuntimeFunctionalityProvider().NUM(" 12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithTrailingWhitespace()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
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

            [Fact]
            public void DateAndTime()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                Assert.Equal(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    GetDefaultRuntimeFunctionalityProvider().NUM(date)
                );
            }

            [Fact]
            public void BytesWithAnInteger()
            {
                // In a loop "FOR i = CBYTE(1) TO CBYTE(5) STEP 1", the "Integer" step of 1 (this would happen with an implicit step too, since that defaults
                // to an "Integer" 1), the loop variable will be "Integer" since it must be a type that can contain all of the constraints. In order to have
                // a loop variable of type "Byte" the loop would need to be of the form "FOR i = CBYTE(1) TO CBYTE(5) STEP CBYTE(1)".
                Assert.Equal(
                    (Int16)1,
                    GetDefaultRuntimeFunctionalityProvider().NUM((byte)1, (byte)5, (Int16)1)
                );
            }

            [Fact]
            public void DateWithDoublesThatAreWithinDateAcceptableRange()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                Assert.Equal(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    GetDefaultRuntimeFunctionalityProvider().NUM(date, 1d)
                );
            }

            [Fact]
            public void DateWithDoublesThatAreNotWithinDateAcceptableRange()
            {
                Assert.Throws<OverflowException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().NUM(new DateTime(2015, 1, 25, 17, 16, 0), double.MaxValue);
                });
            }

            // TODO: String with underscores, hashes, exclamation marks

            // TODO: Negatives
            // TODO: Multiple negatives

            // TODO: Add some tests with numericValuesTheTypeMustBeAbleToContain values, particularly around dates (and decimals)
        }
    }
}
