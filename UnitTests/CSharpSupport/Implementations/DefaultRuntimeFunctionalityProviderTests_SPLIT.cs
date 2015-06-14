using System;
using System.Collections.Generic;
using System.Globalization;
using CSharpSupport;
using CSharpSupport.Exceptions;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class SPLIT : CultureOverridingTests // There is date formatting involved in one of the tests so don't leave culture to chance
        {
            public SPLIT() : base(new CultureInfo("en-GB")) { }

            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object delimiter, object[] expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().SPLIT(value, delimiter));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value, object delimiter)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SPLIT(value, delimiter);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value, object delimiter)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SPLIT(value, delimiter);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value, object delimiter)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SPLIT(value, delimiter);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Empty and blank string get special treatment - rather than returning an array with a single value (or Empty of blank string, resp) VBScript actually returns an array with zero elements
                    yield return new object[] { "Empty value with ' ' delimiter", null, " ", new object[0] };
                    yield return new object[] { "Blank string value with ' ' delimiter", "", " ", new object[0] };

                    yield return new object[] { "Non-blank string value (\"abc\") with ' ' delimiter", "abc", " ", new object[] { "abc" } };
                    yield return new object[] { "Integer value (1) with ' ' delimiter", 1, " ", new object[] { "1" } };
                    yield return new object[] { "Floating point value (1.23) with ' ' delimiter", 1.23, " ", new object[] { "1.23" } };
                    yield return new object[] { "Floating point value (1.23) with '.' delimiter", 1.23, ".", new object[] { "1", "23" } };
                    yield return new object[] { "Date and time with ' ' delimiter", new DateTime(2015, 5, 18, 23, 41, 28), " ", new object[] { "18/05/2015", "23:41:28" } }; // May vary by current culture
                    yield return new object[] {
                        "Object with default property with value \"abc,.def,.ghi\" against delimiter \",.\"",
                        new exampledefaultpropertytype { result = "abc,.def,.ghi" },
                        ",.",
                        new object[] { "abc", "def", "ghi" }
                    };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null string with ' ' delimiter", DBNull.Value, " " };
                    yield return new object[] { "Null delimiter with 'abc' value ", "abc", DBNull.Value };
                    yield return new object[] { "Null delimiter with Nothing value ", VBScriptConstants.Nothing, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "An empty array value with ' ' delimiter", new object[0], " " };
                    yield return new object[] { "An empty array delimiter with \"abc\" value", "abc", new object[0] };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing value with ' ' delimeter", VBScriptConstants.Nothing, " " };
                    yield return new object[] { "Nothing delimiter with 'abc' value ", "abc", VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing delimiter with Null value ", DBNull.Value, VBScriptConstants.Nothing };
                }
            }
        }
    }
}
