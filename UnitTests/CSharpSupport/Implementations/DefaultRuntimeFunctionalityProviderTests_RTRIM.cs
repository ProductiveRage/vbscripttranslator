using System;
using CSharpSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class RTRIM
        {
            [Fact]
            public void EmptyResultsInBlankString()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().RTRIM(null));
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().RTRIM(DBNull.Value));
            }

            [Fact]
            public void DoesNotRemoveTabs()
            {
                Assert.Equal("\tValue\t", DefaultRuntimeSupportClassFactory.Get().RTRIM("\tValue\t"));
            }

            [Fact]
            public void DoesNotRemoveLineReturns()
            {
                Assert.Equal("\nValue\n", DefaultRuntimeSupportClassFactory.Get().RTRIM("\nValue\n"));
            }

            [Fact]
            public void RemovesMultipleTrailingSpaces()
            {
                Assert.Equal("Value", DefaultRuntimeSupportClassFactory.Get().RTRIM("Value   "));
            }

            [Fact]
            public void RemovesMultipleTrailingButNotLeadingSpaces()
            {
                Assert.Equal("  Value", DefaultRuntimeSupportClassFactory.Get().RTRIM("  Value   "));
            }
        }
    }
}
