using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CSTR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CSTR(value));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CSTR(value);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CSTR(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetExceptionData")]
            public void ObjectVariableNotSetExceptionCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CHR(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, "" };
                    yield return new object[] { "Blank string", "", "" };
                    yield return new object[] { "Populated string", "abc", "abc" };
                    yield return new object[] { "Integer 1", 1, "1" };
                    yield return new object[] { "Floating point 1.23", 1.23, "1.23" };
                    yield return new object[] { "Date", new DateTime(2015, 5, 18, 23, 41, 28), (new DateTime(2015, 5, 18, 23, 41, 28)).ToString() }; // May vary by current culture
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "An empty array", new object[0] };
                    yield return new object[] { "Object with default property which is an empty array", new exampledefaultpropertytype { result = new object[0] } };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetExceptionData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }
        }
    }
}
