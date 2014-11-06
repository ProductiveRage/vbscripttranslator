using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndForTranslationTests
    {
        [Fact]
        public void AscendingLoopWithImplicitStep()
        {
            var source = @"
                Dim i: For i = 1 To 5
                Next
            ";
            var expected = new[]
            {
                "for (_outer.i = 1; _.NUM(_outer.i) < 5; _outer.i = _.NUM(_outer.i) + 1)",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        // TODO: Fixed-ascending with step

        // TODO: Fixed-ascending with negative step (nop)

        // TODO: Fixed-Descending without step (nop)

        // TODO Fixed-Descending with step

        // TODO: Various variable-ascending/descending/step combinations
    }
}
