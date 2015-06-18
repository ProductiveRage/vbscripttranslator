using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using System;
using System.Runtime.InteropServices;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CDBL
        {
            [Fact]
            public void Empty()
            {
                Assert.Equal(
                    0d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(null)
                );
            }

            [Fact]
            public void Null()
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDBL(DBNull.Value);
                });
            }

            [Fact]
            public void BlankString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDBL("");
                });
            }

            [Fact]
            public void NonNumericString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDBL("a");
                });
            }

            [Fact]
            public void PositiveNumberAsString()
            {
                Assert.Equal(
                    123.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL("123.4")
                );
            }

            [Fact]
            public void PositiveNumberAsStringWithLeadingAndTrailingWhitespace()
            {
                Assert.Equal(
                    123.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(" 123.4 ")
                );
            }

            [Fact]
            public void PositiveNumberWithNoZeroBeforeDecimalPoint()
            {
                Assert.Equal(
                    0.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(" .4 ")
                );
            }

            [Fact]
            public void NegativeNumberWithNoZeroBeforeDecimalPoint()
            {
                Assert.Equal(
                    -0.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(" -.4 ")
                );
            }

            [Fact]
            public void NegativeNumberWithNoZeroBeforeDecimalPointAndSpaceBetweenSignAndPoint()
            {
                Assert.Equal(
                    -0.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(" - .4 ")
                );
            }

            [Fact]
            public void NegativeNumberAsString()
            {
                Assert.Equal(
                    -123.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL("-123.4")
                );
            }

            [Fact]
            public void Nothing()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDBL(nothing);
                });
            }

            [Fact]
            public void ObjectWithoutDefaultProperty()
            {
                Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDBL(new object());
                });
            }

            [Fact]
            public void ObjectWithDefaultProperty()
            {
                var target = new exampledefaultpropertytype { result = 123.4 };
                Assert.Equal(
                    123.4,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(target)
                );
            }

            [Fact]
            public void Zero()
            {
                Assert.Equal(
                    0d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(0)
                );
            }

            [Fact]
            public void PlusOne()
            {
                Assert.Equal(
                    1d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(1)
                );
            }

            [Fact]
            public void MinusOne()
            {
                Assert.Equal(
                    -1d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(-1)
                );
            }

            [Fact]
            public void OnePointOne()
            {
                Assert.Equal(
                    1.1d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(1.1)
                );
            }

            [Fact]
            public void DateAndTime()
            {
                Assert.Equal(
                    42026.8410300926d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(new DateTime(2015, 1, 22, 20, 11, 5, 0)),
                    10 // This test fails without specifying precision
                );
            }

            [Fact]
            public void True()
            {
                Assert.Equal(
                    -1d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(true)
                );
            }

            [Fact]
            public void False()
            {
                Assert.Equal(
                    0d,
                    DefaultRuntimeSupportClassFactory.Get().CDBL(false)
                );
            }
        }
    }
}
