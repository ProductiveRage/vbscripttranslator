using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CCUR
        {
            [Fact]
            public void JustBeforePositiveOverflow()
            {
                Assert.Equal(
                    VBScriptConstants.MaxCurrencyValue,
                    DefaultRuntimeSupportClassFactory.Get().CCUR(VBScriptConstants.MaxCurrencyValue)
                );
            }

            [Fact]
            public void PositiveOverflow()
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CCUR(VBScriptConstants.MaxCurrencyValue + 0.000001m);
                });
            }

            [Fact]
            public void JustBeforeNegativeOverflow()
            {
                Assert.Equal(
                    VBScriptConstants.MinCurrencyValue,
                    DefaultRuntimeSupportClassFactory.Get().CCUR(VBScriptConstants.MinCurrencyValue)
                );
            }

            [Fact]
            public void NegativeOverflow()
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CCUR(VBScriptConstants.MinCurrencyValue - 0.000001m);
                });
            }
        }
    }
}
