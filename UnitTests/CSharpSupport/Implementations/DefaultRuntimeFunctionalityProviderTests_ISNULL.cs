using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ISNULL
        {
            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object value)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().ISNULL(value));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object value)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().ISNULL(value));
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Zero", 0 };
                    yield return new object[] { "Blank string", "" };
                    yield return new object[] { "Unintialised array", new object[0] };
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
                }
            }
        }
    }
}
