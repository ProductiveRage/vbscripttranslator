using CSharpSupport;
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
                    DefaultRuntimeSupportClassFactory.Get().NUM(null)
                );
            }

            [Fact]
            public void Null()
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM(DBNull.Value);
                });
            }

            [Fact]
            public void True()
            {
                Assert.Equal(
                    (Int16)(-1),
                    DefaultRuntimeSupportClassFactory.Get().NUM(true)
                );
            }

            [Fact]
            public void False()
            {
                Assert.Equal(
                    (Int16)0,
                    DefaultRuntimeSupportClassFactory.Get().NUM(false)
                );
            }

            [Fact]
            public void BlankString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM("");
                });
            }

            [Fact]
            public void PositiveIntegerString()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Get().NUM("12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithLeadingWhitespace()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Get().NUM(" 12")
                );
            }

            [Fact]
            public void PositiveIntegerStringWithTrailingWhitespace()
            {
                Assert.Equal(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Get().NUM("12 ")
                );
            }

            [Fact]
            public void TwoIntegersSeparatedByWhitespace()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM("1 1");
                });
            }

            [Fact]
            public void PositiveDecimalString()
            {
                Assert.Equal(
                    1.2,
                    DefaultRuntimeSupportClassFactory.Get().NUM("1.2")
                );
            }

            [Fact]
            public void PseudoNumberWithMultipleDecimalPoints()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM("1.1.0");
                });
            }

            [Fact]
            public void DateAndTime()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                Assert.Equal(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    DefaultRuntimeSupportClassFactory.Get().NUM(date)
                );
            }

            [Fact]
            public void IntegerWithDate()
            {
                Assert.Equal(
                    new DateTime(1899, 12, 31), // This is the VBScript "ZeroDate" plus one day (which is what 1 is translated into in order to become a date)
                    DefaultRuntimeSupportClassFactory.Get().NUM(1, new DateTime(2015, 1, 22, 20, 11, 5, 0))
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
                    DefaultRuntimeSupportClassFactory.Get().NUM((byte)1, (byte)5, (Int16)1)
                );
            }

            [Fact]
            public void DateWithDoublesThatAreWithinDateAcceptableRange()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                Assert.Equal(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    DefaultRuntimeSupportClassFactory.Get().NUM(date, 1d)
                );
            }

            [Fact]
            public void DateWithDoublesThatAreNotWithinDateAcceptableRange()
            {
                Assert.Throws<OverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM(new DateTime(2015, 1, 25, 17, 16, 0), double.MaxValue);
                });
            }

            [Fact]
            public void IntegerWithIntegerValueAsString()
            {
                // Strings are always parsed into doubles, regardless of the size of the value they represent
                Assert.Equal(
                    1d,
                    DefaultRuntimeSupportClassFactory.Get().NUM((Int16)1, "2")
                );
            }

            [Fact]
            public void StringRepresentationsOfDatesAreNotParsed()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM("1/1/2015");
                });
            }

            [Fact]
            public void StringRepresentationsOfISODatesAreNotParsed()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM("2015-01-01");
                });
            }

            [Fact]
            public void DecimalIsEvenBiggerThanDouble()
            {
                // Although the double type can contain a greater range of values than decimal, VBScript prefers decimal if both are present
                Assert.Equal(
                    1m,
                    DefaultRuntimeSupportClassFactory.Get().NUM(1m, 2d)
                );
            }

            [Fact]
            public void DecimalWithDoublesThatAreNotWithinVBScriptCurrencyAcceptableRange()
            {
                // See https://msdn.microsoft.com/en-us/library/9e7a57cf%28v=vs.84%29.aspx for limits of the VBScript data types
                Assert.Throws<OverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NUM(
                        922337203685475m, // Toward the top end of the Currency limit
                        1000000000000000d // Definitely past it
                    );
                });
            }

            [Fact]
            public void IntegerWithDecimal()
            {
                Assert.Equal(
                    1m,
                    DefaultRuntimeSupportClassFactory.Get().NUM(1, 2m)
                );
            }

            // TODO: String with underscores, hashes, exclamation marks

            // TODO: Negatives
            // TODO: Multiple negatives
        }
    }
}
