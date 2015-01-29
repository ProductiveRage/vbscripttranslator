using CSharpSupport;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class NullableNUM
        {
            [Fact]
            public void NullToNumber()
            {
                Assert.Equal(
                    DBNull.Value,
                    DefaultRuntimeSupportClassFactory.Get().NullableNUM(DBNull.Value)
                );
            }
        }
    }
}
