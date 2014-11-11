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
                "for (_outer.i = 1; _.NUM(_outer.i) <= 5; _outer.i = _.NUM(_outer.i) + 1)",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void DescendingLoopWithoutExplicitStepIsOptimisedOut()
        {
            // If the loop range is in the opposite direction to step then it will never be entered in VBScript and so there's no pointing emitting any C# code (this
            // can only be done if the loop start, end and step are known at compile time - here the start and end are numeric and the loop is implicitly one)
            var source = @"
                Dim i: For i = 5 To 1
                Next
            ";
            Assert.Equal(
                new string[0],
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void ZeroStepResultsInInfiniteLoopWhenAscending()
        {
            var source = @"
                Dim i: For i = 1 To 5 Step 0
                Next
            ";
            var expected = new[]
            {
                "for (_outer.i = 1; _.NUM(_outer.i) <= 5;)",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        // TODO public void ZeroStepIsOptimisedOutForDescendingLoop()



        // TODO: Fixed-ascending with step

        // TODO: Fixed-ascending with negative step (nop)

        // TODO: Fixed-Descending without step (nop)

        // TODO Fixed-Descending with step

        // TODO: Various variable-ascending/descending/step combinations
    }
}
