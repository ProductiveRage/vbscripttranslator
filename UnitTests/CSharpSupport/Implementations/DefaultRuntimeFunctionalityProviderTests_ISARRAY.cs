using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Attributes;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ISARRAY
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, bool expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().ISARRAY(value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, false };
                    yield return new object[] { "Null", DBNull.Value, false };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing, false };
                    yield return new object[] { "Zero", 0, false };
                    yield return new object[] { "Blank string", "", false };

                    yield return new object[] { "Empty 1D array", new object[0], true };    // In VBScript: Either "Array()" or "Array(-1)"
                    yield return new object[] { "Empty 2D array", new object[0, 0], true }; // In VBScript: "Array(-1, -1)"
                    yield return new object[] { "Populated 1D array", new object[] { 1 }, true };
                    
                    yield return new object[] { "Object with default property which is Populated 1D array", new exampledefaultpropertytype { result = new object[] { 1 } }, true };
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
