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
                    GetDefaultRuntimeFunctionalityProvider().CDBL(null)
                );
            }

            [Fact]
            public void Null()
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().CDBL(DBNull.Value);
                });
            }

            [Fact]
            public void BlankString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().CDBL("");
                });
            }

            [Fact]
            public void NonNumericString()
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().CDBL("a");
                });
            }

            [Fact]
            public void PositiveNumberAsString()
            {
                Assert.Equal(
                    123.4,
                    GetDefaultRuntimeFunctionalityProvider().CDBL("123.4")
                );
            }

            [Fact]
            public void PositiveNumberAsStringWithLeadingAndTrailingWhitespace()
            {
                Assert.Equal(
                    123.4,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(" 123.4 ")
                );
            }

            [Fact]
            public void NegativeNumberAsString()
            {
                Assert.Equal(
                    -123.4,
                    GetDefaultRuntimeFunctionalityProvider().CDBL("-123.4")
                );
            }

            [Fact]
            public void Nothing()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().CDBL(nothing);
                });
            }

            [Fact]
            public void ObjectWithoutDefaultProperty()
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().CDBL(new object());
                });
            }

            [Fact]
            public void ObjectWithDefaultProperty()
            {
                var target = new exampledefaultpropertytype { result = 123.4 };
                Assert.Equal(
                    123.4,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(target)
                );
            }

            [Fact]
            public void Zero()
            {
                Assert.Equal(
                    0d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(0)
                );
            }

            [Fact]
            public void PlusOne()
            {
                Assert.Equal(
                    1d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(1)
                );
            }

            [Fact]
            public void MinusOne()
            {
                Assert.Equal(
                    -1d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(-1)
                );
            }

            [Fact]
            public void OnePointOne()
            {
                Assert.Equal(
                    1.1d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(1.1)
                );
            }

            [Fact]
            public void DateAndTime()
            {
                Assert.Equal(
                    42026.8410300926d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(new DateTime(2015, 1, 22, 20, 11, 5, 0)),
                    10 // This test fails without specifying precision
                );
            }

            [Fact]
            public void True()
            {
                Assert.Equal(
                    -1d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(true)
                );
            }

            [Fact]
            public void False()
            {
                Assert.Equal(
                    0d,
                    GetDefaultRuntimeFunctionalityProvider().CDBL(false)
                );
            }

            /// <summary>
            /// This is an example of the type of class that may be emitted by the translation process, one with a parameter-less default member
            /// </summary>
            [TranslatedProperty("ExampleDefaultPropertyType")]
            private class exampledefaultpropertytype
            {
                [IsDefault]
                public object result { get; set; }
            }
        }
    }
}
