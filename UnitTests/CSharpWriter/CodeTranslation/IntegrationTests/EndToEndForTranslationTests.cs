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

        /// <summary>
        /// If the loop range is in the opposite direction to step then it will never be entered in VBScript and so there's no pointing emitting any C# code (this
        /// can only be done if the loop start, end and step are known at compile time - here the start and end are numeric and the loop is implicitly one)
        /// </summary>
        [Fact]
        public void DescendingLoopWithoutExplicitStepIsOptimisedOut()
        {
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
        public void DescendingLoopWithExplicitNegativeStep()
        {
            var source = @"
                Dim i: For i = 5 To 1 Step -1
                Next
            ";
            var expected = new[]
            {
                "for (_outer.i = 5; _.NUM(_outer.i) >= 1; _outer.i = _.NUM(_outer.i) - 1)",
                "{",
                "}"
            };
            Assert.Equal(
                expected,
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

        /// <summary>
        /// If the loop has fixed contraints that indicate a negative direction and a zero step, the loop will not be entered and can be optimised out
        /// </summary>
        [Fact]
        public void ZeroStepIsOptimisedOutForDescendingLoop()
        {
            var source = @"
                Dim i: For i = 5 To 1 Step 0
                Next
            ";
            Assert.Equal(
                new string[0],
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop has fixed contraints that indicate a negative direction and a zero step, the loop will not be entered and can be optimised out
        /// </summary>
        [Fact]
        public void FixedNegativeStepResultsInLoopBeingOptimisedOutIfItIsFixedAndPositive()
        {
            var source = @"
                Dim i: For i = 1 To 5 Step - 1
                Next
            ";
            Assert.Equal(
                new string[0],
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a loop is known at compile time to run in a negative direction and no step is specified then the loop is never entered and can be optimised out
        /// </summary>
        [Fact]
        public void FixedNegativeLoopWithoutExplicitStepIsOptimisedOut()
        {
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
        public void FixedAscendingLoopWithExplicitPositiveStep()
        {
            var source = @"
                Dim i: For i = 1 To 5 Step 2
                Next
            ";
            var expected = new[]
            {
                "for (_outer.i = 1; _.NUM(_outer.i) <= 5; _outer.i = _.NUM(_outer.i) + 2)",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void FixedDescendingLoopWithExplicitNegativeStep()
        {
            var source = @"
                Dim i: For i = 5 To 1 Step -1
                Next
            ";
            var expected = new[]
            {
                "for (_outer.i = 5; _.NUM(_outer.i) >= 1; _outer.i = _.NUM(_outer.i) - 1)",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        // TODO: Various variable-ascending/descending/step combinations
    }
}
