using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class UBOUND
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, int dimension, int expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().UBOUND(value, dimension));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value, int dimension)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().UBOUND(value, dimension);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value, int dimension)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().UBOUND(value, dimension);
                });
            }

            [Theory, MemberData("SubscriptOutOfRangeData")]
            public void SubscriptOutOfRangeCases(string description, object value, int dimension)
            {
                Assert.Throws<SubscriptOutOfRangeException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().UBOUND(value, dimension);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty 1D array", new object[0], 1, -1 }; // In VBScript: Either "Array()" or "Dim arr: ReDim arr(-1)"
                    yield return new object[] { "1D array with a single item", new object[1], 1, 0 };
                    yield return new object[] { "Object with default property which is Populated 1D array", new exampledefaultpropertytype { result = new object[] { 1 } }, 1, 0 };

                    yield return new object[] { "2D array where first dimension is larger and first dimension is requested", new exampledefaultpropertytype { result = new object[7, 2] }, 1, 6 };
                    yield return new object[] { "2D array where first dimension is larger and second dimension is requested", new exampledefaultpropertytype { result = new object[7, 2] }, 2, 1 };
                    yield return new object[] { "2D array where second dimension is larger and first dimension is requested", new exampledefaultpropertytype { result = new object[2, 7] }, 1, 1 };
                    yield return new object[] { "2D array where second dimension is larger and second dimension is requested", new exampledefaultpropertytype { result = new object[2, 7] }, 2, 6 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Empty", null, 1 };
                    yield return new object[] { "Null", DBNull.Value, 1 };
                    yield return new object[] { "Blank string", "", 1 };
                    yield return new object[] { "Object with default property which is Emty", new exampledefaultpropertytype(), 1 };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing, 1 };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing }, 1 };
                }
            }

            public static IEnumerable<object[]> SubscriptOutOfRangeData
            {
                get
                {
                    yield return new object[] { "1D array where dimension 2 is requested", new object[1], 2 };
                }
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
