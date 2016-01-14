using System;
using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class LTRIM
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
            public void RemovesMultipleLeadingSpaces()
            {
                Assert.Equal("Value", DefaultRuntimeSupportClassFactory.Get().LTRIM("  Value"));
            }

            [Fact]
            public void RemovesMultipleLeadingButNotTrailingSpaces()
            {
                Assert.Equal("Value   ", DefaultRuntimeSupportClassFactory.Get().LTRIM("  Value   "));
            }
        }
    }
}
