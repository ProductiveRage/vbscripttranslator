using System;
using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class LCASE
        {
            [Fact]
            public void EmptyResultsInBlankString()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().LCASE(null));
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().LCASE(DBNull.Value));
            }

            [Fact]
            public void Test()
            {
                Assert.Equal("test", DefaultRuntimeSupportClassFactory.Get().LCASE("Test"));
            }
        }
    }
}
