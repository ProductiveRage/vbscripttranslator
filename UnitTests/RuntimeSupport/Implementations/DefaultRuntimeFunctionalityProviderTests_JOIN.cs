using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class JOIN
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object delimiter, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().JOIN(value, delimiter));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value, object delimiter)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().JOIN(value, delimiter);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value, object delimiter)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().JOIN(value, delimiter);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value, object delimiter)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().JOIN(value, delimiter);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Uninitialised object array", new object[0], " ", "" };
                    yield return new object[] { "Uninitialised object array with Empty delimiter", new object[0], null, "" };

                    yield return new object[] { "1D object array of numeric values with comma delimiter", new object[] { 1, 2, 3 }, ",", "1,2,3" };
                    yield return new object[] { "1D object array of numeric values with comma+space delimiter", new object[] { 1, 2, 3 }, ", ", "1, 2, 3" };
                    yield return new object[] { "1D object array of numeric/Empty values with comma delimiter", new object[] { 1, null, 3 }, ",", "1,,3" };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value, " " };
                    yield return new object[] { "Null delimiter", new object[0], DBNull.Value };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, " " };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Empty", null, " " };
                    yield return new object[] { "Zero", 0, " " };
                    yield return new object[] { "Blank string", "", " " };
                    yield return new object[] { "String: \"Test\"", "Test", " " };
                    yield return new object[] { "2D object array", new object[0, 0], " " };
                    yield return new object[] { "1D object array of numeric/Null values with comma delimiter", new object[] { 1, DBNull.Value, 3 }, "," }; // Would have expected invalid-use-of-null! But VBScript goes for type-mismatch..
                    yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" }, " " };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing, " " };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing }, " " };
                    yield return new object[] { "1D object array of numeric/Nothing values with comma delimiter", new object[] { 1, VBScriptConstants.Nothing, 3 }, "," };
                }
            }
        }
    }
}
