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
                "foreach (_env.value in _.ENUMERABLE(_env.values))",
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
        /// Error-trapping around For Each loops has the curious behaviour that if it is not possible to enumerate over the target (eg. attempting to enumerate
        /// over a string, which is not supported by VBScript) or if there is an error in evaluating the target (eg. attempting to enumerate over the return
        /// value of a function where that function raises an error) then the loop will be entered once, but the loop variable's value will not be altered
        /// (so if it had no value before the loop was reached, it will continue to have no value - but if it already had a value before the loop, this
        /// value will remain unaltered). If error-trapping may be enabled, then the translated code tries to define a set to enumerate over, leaving
        /// it null if this fails (and if error-trapping is actually enabled; if it isn't then the error will be raised and the scope exited). When
        /// the loop is reached, if the target set is null then it falls back to a single-element array whose item is the loop variable's original
        /// value, which emulates VBScript's behaviour. If there is no ON ERROR RESUME NEXT around the code then none of this extra logic appears
        /// in the output (see the SimpleCaseWithoutErrorHandling test), but if error-trapping may be enabled then this code must be present
        /// (even if there's a chance that an ON ERROR GOTO 0 will disable the error-trapping in the current scope at some point at runtime).
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
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "IEnumerable enumerationContent2 = null;",
                "_.HANDLEERROR(errOn1, () => {",
                "    enumerationContent2 = _.ENUMERABLE(_env.values)",
                "}",
                "foreach (_env.value in enumerationContent2 ?? new object[] { _env.value })",
                "{",
                "    _.HANDLEERROR(errOn1, () => {",
                "        _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.value, v3 => { _env.value = v3; }));",
                "    });",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);",

            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is extremely similar to SimpleCaseWithErrorHandling but it demonstrates some of the capabilities and limitations of the static analysis
        /// regarding potential error-trapping scenarios. The presence of the ON ERROR RESUME NEXT means that the translation process has to consider that
        /// some error-trapping may occur within the current scope. However, it is not intelligent enough to realise that it is then immediately disabled
        /// and so the additional error-trapping-related code around the FOR EACH loop is not necessary. It IS intelligent enough to realise that there
        /// is no scenario where error-trapping could be enabled within the loop, so the WScript.Echo call is not wrapped in a HANDLEERROR call. This
        /// could be improved in the future but I thought it was worth having this test to illustrate the discrepancy.
        /// </summary>
        [Fact]
        public void SimpleCaseWithErrorHandlingButWithItDisabledAtRunTimeBeforeTheLoop()
        {
            var source = @"
                On Error Resume Next
                On Error Goto 0
                For Each value In values
                    WScript.Echo value
                Next
            ";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "_.STOPERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "IEnumerable enumerationContent2 = null;",
                "_.HANDLEERROR(errOn1, () => {",
                "    enumerationContent2 = _.ENUMERABLE(_env.values)",
                "}",
                "foreach (_env.value in enumerationContent2 ?? new object[] { _env.value })",
                "{",
                "    _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.value, v3 => { _env.value = v3; }));",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);",

            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
