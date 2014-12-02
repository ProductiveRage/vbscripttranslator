using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndForEachTranslationTests
    {
        [Fact]
        public void SimpleCaseWithoutErrorHandling()
        {
            var source = @"
                For Each value IN values
                    WScript.Echo value
                Next
            ";
            var expected = new[]
            {
                "for each (_env.value in _.ENUMERABLE(_env.values)",
                "{",
                "_.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.value, v1 => { _env.value = v1; }));",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
