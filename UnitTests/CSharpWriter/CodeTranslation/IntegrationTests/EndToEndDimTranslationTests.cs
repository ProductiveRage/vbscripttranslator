using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndDimTranslationTests
	{
		[Fact]
		public void DimInsideFunction()
		{
			var source = @"
				Function F1()
					Dim myVariable
				End Function
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"object retVal1 = null;",
				"object myvariable = null;",
				"return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		[Fact]
		public void DimWithDimensionsInsideFunction()
		{
			var source = @"
				Function F1()
					Dim myArray(63)
				End Function
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"object retVal1 = null;",
				"object myarray = new object[64];",
				"return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}
	}
}
