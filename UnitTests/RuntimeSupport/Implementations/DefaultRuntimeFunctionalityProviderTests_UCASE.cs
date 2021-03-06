﻿using System;
using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class UCASE
        {
            [Fact]
            public void EmptyResultsInBlankString()
            {
                Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().UCASE(null));
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().UCASE(DBNull.Value));
            }

            [Fact]
            public void Test()
            {
                Assert.Equal("TEST", DefaultRuntimeSupportClassFactory.Get().UCASE("Test"));
            }
        }
    }
}
