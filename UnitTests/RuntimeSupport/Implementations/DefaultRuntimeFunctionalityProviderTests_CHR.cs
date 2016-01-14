using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CHR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, Char expectedResult)
            {
                Assert.Equal(new string(expectedResult, 1), DefaultRuntimeSupportClassFactory.Get().CHR(value));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CHR(value);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CHR(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CHR(value);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object value)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CHR(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, (char)0 };
                    yield return new object[] { "0", 0, (char)0 };
                    yield return new object[] { "1", 1, (char)1 };
                    yield return new object[] { "255", 255, (char)255 };
                    yield return new object[] { "255.4", 255.4, (char)255 };
                    yield return new object[] { "-0.5", -0.5, (char)0 };
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
                    yield return new object[] { "Blank string", ""};
                    yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" } };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "255.5", 255.5 };
                    yield return new object[] { "-0.6", -0.6 };
                }
            }
        }
    }
}
