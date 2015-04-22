using System;
using CSharpSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class TRIM
        {
            [Fact]
            public void EmptyResultsInBlankString()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().LTRIM(null));
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().LTRIM(DBNull.Value));
            }

            [Fact]
            public void DoesNotRemoveTabs()
            {
                Assert.Equal("\tValue\t", DefaultRuntimeSupportClassFactory.Get().LTRIM("\tValue\t"));
            }

            [Fact]
            public void DoesNotRemoveLineReturns()
            {
                Assert.Equal("\nValue\n", DefaultRuntimeSupportClassFactory.Get().LTRIM("\nValue\n"));
            }

            [Fact]
            public void RemovesMultipleLeadingAndTrailingSpaces()
            {
                Assert.Equal("  Value   ", DefaultRuntimeSupportClassFactory.Get().LTRIM("Value"));
            }
        }
    }
}
