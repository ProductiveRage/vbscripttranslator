using System;
using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class ESCAPE
		{
			[Fact]
			public void EmptyResultsInBlankString()
			{
				Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().ESCAPE(null));
			}

			[Fact]
			public void NullResultsInNull()
			{
				Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().ESCAPE(DBNull.Value));
			}

			[Fact]
			public void PlainString()
			{
				Assert.Equal("test", DefaultRuntimeSupportClassFactory.Get().ESCAPE("test"));
			}

			[Fact]
			public void ComplexString()
			{
				Assert.Equal("%22T%FCst%20the%2Cth+in%252Bg%20%u0107%22", DefaultRuntimeSupportClassFactory.Get().ESCAPE("\"Tüst the,th+in%2Bg ć\""));
			}

			[Fact]
			public void NonEscapedCharacters()
			{
				Assert.Equal("@*_+-./", DefaultRuntimeSupportClassFactory.Get().ESCAPE("@*_+-./"));
			}
		}
	}
}
