using System;
using CSharpSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class LEN
        {
            [Fact]
            public void EmptyResultsInZero()
            {
                Assert.Equal((int)0, DefaultRuntimeSupportClassFactory.Get().LEN(null)); // This should return an int ("Long" in VBScript parlance)
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().LEN(DBNull.Value));
            }

            [Fact]
            public void Test()
            {
                Assert.Equal(4, DefaultRuntimeSupportClassFactory.Get().LEN("Test"));
            }
        }
    }
}
