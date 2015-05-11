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
            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object value)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().ISARRAY(value));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object value)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().ISARRAY(value));
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "Empty 1D array", new object[0] };    // In VBScript: Either "Array()" or "Dim arr: ReDim arr(-1)"
                    yield return new object[] { "Empty 2D array", new object[0, 0] }; // In VBScript: "Dim arr: ReDim arr(-1, -1)"
                    yield return new object[] { "Populated 1D array", new object[] { 1 } };
                    yield return new object[] { "Object with default property which is Populated 1D array", new exampledefaultpropertytype { result = new object[] { 1 } } };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Zero", 0 };
                    yield return new object[] { "Blank string", "" };
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
