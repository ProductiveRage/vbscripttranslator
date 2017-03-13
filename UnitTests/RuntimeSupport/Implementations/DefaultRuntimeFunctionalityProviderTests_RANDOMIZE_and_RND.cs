using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class RANDOMIZE_and_RND
		{
			[Fact]
			public void RandomizeSeedReturnsConsistentValuesFirstTimeItIsUsedForRuntimeSupportFactoryInstance()
			{
				const int seed = 123;
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE(seed);
					value1 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE(seed);
					value2 = _.RND();
				}
				Assert.Equal(value1, value2);
			}

			[Fact]
			public void RandomizeSeedReturnsDifferentSequencesIfUsedWithinSameRuntimeSupportFactoryInstance()
			{
				const int seed = 123;
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE(seed);
					value1 = _.RND();

					_.RANDOMIZE(seed);
					value2 = _.RND();
				}
				Assert.NotEqual(value1, value2);
			}

			[Fact]
			public void CallingRndWithZeroReturnsPreviousNumber()
			{
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					value1 = _.RND();
					value2 = _.RND(0);
				}
				Assert.Equal(value1, value2);
			}

			[Fact]
			public void CallingRndWithNegativeValueSetsSeedToThatValue()
			{
				const int seed = -123;
				float[] values1, values2;
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					values1 = new[] { _.RND(seed), _.RND(), _.RND() };
					values2 = new[] { _.RND(seed), _.RND(), _.RND() };
				}
				Assert.Equal(values1, values2);
			}

			/// <summary>
			/// The precision of the RANDOMIZE seed is limited to a Single (in VBScript parlance, which I think is equivalent to .NET) - so the extra digit on 1.1111111 does not make any
			/// difference compared to 1.111111 (though going one smaller at 1.11111 WILL result in a different sequence being generated)
			/// </summary>
			[Fact]
			public void RandomizeSeedValueHasLimitedPrecision()
			{
				// The values
				//  1.111111
				// and
				//  1.1111111
				// will result in the same random number streams being generated
				float value1, value2, value3;
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE("1.111111");
					value1 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE("1.1111111");
					value2 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Get())
				{
					_.RANDOMIZE("1.11111");
					value3 = _.RND();
				}
				Assert.Equal(value1, value2);
				Assert.NotEqual(value1, value3);
			}
		}
	}
}
