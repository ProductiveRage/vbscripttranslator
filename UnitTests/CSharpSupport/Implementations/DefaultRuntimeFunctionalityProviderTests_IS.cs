using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class IS
        {
            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object l, object r)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().IS(l, r));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object l, object r)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().IS(l, r));
            }

            [Theory, MemberData("ObjectRequiredData")]
            public void ObjectRequiredCases(string description, object l, object r)
            {
                Assert.Throws<ObjectRequiredException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().IS(l, r);
                });
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "Nothing vs Nothing", VBScriptConstants.Nothing, VBScriptConstants.Nothing };
                    var x = new exampledefaultpropertytype();
                    yield return new object[] { "ClassInstance vs SameClassInstance", x, x };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    var x = new exampledefaultpropertytype();
                    yield return new object[] { "ClassInstance vs Nothing", x, VBScriptConstants.Nothing };
                    var y = new exampledefaultpropertytype();
                    yield return new object[] { "ClassInstance vs DifferentClassInstance", x, y };
                }
            }

            public static IEnumerable<object[]> ObjectRequiredData
            {
                get
                {
                    yield return new object[] { "Empty vs Nothing", null, VBScriptConstants.Nothing };
                    yield return new object[] { "Null vs Nothing", DBNull.Value, VBScriptConstants.Nothing };
                    yield return new object[] { "Zero vs Nothing", 0, VBScriptConstants.Nothing };
                    yield return new object[] { "Blank string vs Nothing", "", VBScriptConstants.Nothing };
                    yield return new object[] { "Date vs Nothing", new DateTime(2015, 5, 28, 16, 4, 58), VBScriptConstants.Nothing };
                    yield return new object[] { "Unintialised array vs Nothing", new object[0], VBScriptConstants.Nothing };
                }
            }
        }
    }
}
