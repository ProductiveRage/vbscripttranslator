using System;
using CSharpSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ARRAY
        {
            /// <summary>
            /// The ARRAY method should never be called with a null values array - if it is called with zero arguments then the array should be a zero-element array instance, not null
            /// </summary>
            [Fact]
            public void Null()
            {
                Assert.Throws<ArgumentNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().ARRAY(null);
                });
            }

            [Fact]
            public void ZeroElements()
            {
                Assert.Equal(new object[0], DefaultRuntimeSupportClassFactory.Get().ARRAY());
            }

            [Fact]
            public void OneElement()
            {
                Assert.Equal(new object[] { 1 }, DefaultRuntimeSupportClassFactory.Get().ARRAY(1));
            }

            [Fact]
            public void TwoElements()
            {
                Assert.Equal(new object[] { 1, 2 }, DefaultRuntimeSupportClassFactory.Get().ARRAY(1, 2));
            }
        }
    }
}
