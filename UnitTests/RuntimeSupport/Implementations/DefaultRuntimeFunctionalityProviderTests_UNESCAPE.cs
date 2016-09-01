using System;
using VBScriptTranslator.RuntimeSupport;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class UNESCAPE
		{
			[Fact]
			public void EmptyResultsInBlankString()
			{
				Assert.Equal("", DefaultRuntimeSupportClassFactory.Get().UNESCAPE(null));
			}

			[Fact]
			public void NullResultsInNull()
			{
				Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().UNESCAPE(DBNull.Value));
			}

			[Fact]
			public void PlainString()
			{
				Assert.Equal("test", DefaultRuntimeSupportClassFactory.Get().UNESCAPE("test"));
			}

			[Fact]
			public void ComplexString()
			{
				Assert.Equal("\"Tüst the,th+in%2Bg ć\"", DefaultRuntimeSupportClassFactory.Get().UNESCAPE("%22T%FCst%20the%2Cth+in%252Bg%20%u0107%22"));
			}

			[Fact]
			public void NonEscapedCharacters()
			{
				Assert.Equal("@*_+-./", DefaultRuntimeSupportClassFactory.Get().UNESCAPE("@*_+-./"));
			}
		}
	}
}
