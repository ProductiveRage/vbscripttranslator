using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndClassTranslationTests
    {
        [Fact]
        public void EndProperty()
        {
            var source = @"
                PUBLIC PROPERTY GET Name
                END PROPERTY
            ";
            var expected = new[]
            {
                "public object name()",
                "{",
                "  return null;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
