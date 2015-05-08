using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Attributes;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ISNULL
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, bool expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().ISNULL(value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value, true };

                    yield return new object[] { "Empty", null, false };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing, false };
                    yield return new object[] { "Zero", 0, false };
                    yield return new object[] { "Blank string", "", false };
                    yield return new object[] { "Unintialised array", new object[0], false };

                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), false };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, true };
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
