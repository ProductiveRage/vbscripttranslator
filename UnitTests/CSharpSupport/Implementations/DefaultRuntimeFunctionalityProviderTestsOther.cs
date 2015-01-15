using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class NULLABLENUM
        {
            [Fact]
            public void NullToNumber()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().NULLABLENUM(DBNull.Value)
                );
            }
        }
    }
}
