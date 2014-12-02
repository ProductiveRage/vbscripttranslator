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
                For Each value In values
                    WScript.Echo value
                Next
            ";
            var expected = new[]
            {
                "for each (_env.value in _.ENUMERABLE(_env.values)",
                "{",
                "    _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.value, v1 => { _env.value = v1; }));",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// Error-trapping with For Each loops is must less complicated than For loops (where error-trapping is required around the constraints evaluation - if
        /// they're not compile-time constants - and then distinct error-trapping around the loop, if the constraint evaluation didn't fail)
        /// </summary>
        [Fact]
        public void SimpleCaseWithErrorHandling()
        {
            var source = @"
                On Error Resume Next
                For Each value In values
                    WScript.Echo value
                Next
            ";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPING(errOn1);",
                "_.HANDLEERROR(errOn1, () => {",
                "    for each (_env.value in _.ENUMERABLE(_env.values)",
                "    {",
                "        _.HANDLEERROR(errOn1, () => {",
                "            _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.value, v2 => { _env.value = v2; }));",
                "        });",
                "    }",
                "});",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
