using CSharpSupport;
using CSharpSupport.Exceptions;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class RIGHT
        {
            /// <summary>
            /// Passing in VBScript Empty as the string will return in a blank string being returned (so long as the length argument can be interpreted as a non-negative number)
            /// </summary>
            [Fact]
            public void EmptyLengthOneReturnsBlankString()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().RIGHT(null, 1));
            }

            /// <summary>
            /// Passing in VBScript Null as the string will return in VBScript Null being returned (so long as the length argument can be interpreted as a non-negative number)
            /// </summary>
            [Fact]
            public void NullLengthOneReturnsNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().RIGHT(DBNull.Value, 1));
            }

            [Fact]
            public void ZeroLengthIsAcceptable()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().RIGHT("", 0));
            }

            [Fact]
            public void NegativeLengthIsNotAcceptable()
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().RIGHT("", -1);
                });
            }

            [Fact]
            public void EmptyLengthIsTreatedAsZeroLength()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().RIGHT("abc", null));
            }

            [Fact]
            public void MaxLengthLongerThanInputStringLengthIsTreatedAsEqualingInputStringLength()
            {
                Assert.Equal("abc", DefaultRuntimeSupportClassFactory.Get().RIGHT("abc", 10));
            }

            [Fact]
            public void NullLengthIsNotAcceptable()
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().RIGHT("", DBNull.Value);
                });
            }

            [Fact]
            public void EnormousLengthResultsInOverflow()
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().RIGHT("", 1000000000000000);
                });
            }

            // These tests all illustrate that VBScript's standard "banker's rounding" is applied to fractional lengths
            [Fact]
            public void LengthZeroPointFiveTreatedAsLengthZero()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 0.5));
            }
            [Fact]
            public void LengthZeroPointNineTreatedAsLengthOne()
            {
                Assert.Equal("d", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 0.9));
            }
            [Fact]
            public void LengthOnePointFiveTreatedAsLengthTwo()
            {
                Assert.Equal("cd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 1.5));
            }
            [Fact]
            public void LengthOnePointNineTreatedAsLengthTwo()
            {
                Assert.Equal("cd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 1.9));
            }
            [Fact]
            public void LengthTwoPointFiveTreatedAsLengthTwo()
            {
                Assert.Equal("cd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 2.5));
            }
            [Fact]
            public void LengthTwoPointNineTreatedAsLengthThree()
            {
                Assert.Equal("bcd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 2.9));
            }
            [Fact]
            public void LengthThreePointFiveTreatedAsLengthFour()
            {
                Assert.Equal("abcd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 3.5));
            }
            [Fact]
            public void LengthThreePointNineTreatedAsLengthFour()
            {
                Assert.Equal("abcd", DefaultRuntimeSupportClassFactory.Get().RIGHT("abcd", 3.9));
            }
        }
    }
}
