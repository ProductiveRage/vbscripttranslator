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

        /// <summary>
        /// If the loop start, end and step values are not known until runtime then their values must be determined once and then applied to a loop (in
        /// VBScript, the constraints are not re-evaulated each loop iteration). The loop may only be entered if there is a zero or positive step and a
        /// non-descending loop or if there is a negative step and a descending loop. Similarly, the termination condition operator may be a less-than-
        /// or-equal-to comparison or a greater-than-or-equal-to, depending upon loop direction.
        /// </summary>
        [Fact]
        public void RuntimeVariableLoopBoundariesAndStep()
        {
            var source = @"
                For i = a To b Step c
                Next
            ";
            var expected = new[]
            {
                "var loopStart1 = _.NUM(_env.a);",
                "var loopEnd2 = _.NUM(_env.b);",
                "var loopStep3 = _.NUM(_env.c);",
                "if (((loopStart1 <= loopEnd2) && (loopStep3 >= 0)) || ((loopStart1 > loopEnd2) && (loopStep3 < 0)))",
                "{",
                "for (_env.i = loopStart1; ((loopStep3 >= 0) && (_.NUM(_env.i) <= loopEnd2)) || ((loopStep3 < 0) && (_.NUM(_env.i) >= loopEnd2)); _env.i = _.NUM(_env.i) + loopStep3)",
                "{",
                "}",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If there are non-compile-time-known-numeric-constant constraints and error-trapping may be enabled, then the constraints must be evaluated
        /// first, and no further work undertaken if this fails. Then there must be error-trapping around the loop itself, so that if the termination
        /// condition or loop-variable-addition/subtraction fails then the loop will terminate. Then there must be error-trapping around each statement
        /// within the loop, so that if any one of them fails then the others may still be processed (if the error-trapping token is enabled at that
        /// point during the runtime execution).
        /// </summary>
        [Fact]
        public void RuntimeVariableLoopBoundariesWithErrorTrapping()
        {
            var source = @"
                On Error Resume Next
                For i = a To b
                    WScript.Echo i
                Next
            ";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPING(errOn1);",
                "double loopStart2 = 0, loopEnd3 = 0;",
                "var loopConstraintsInitialised4 = false;",
                "_.HANDLEERROR(errOn1, () =>",
                "{",
                "    loopStart2 = _.NUM(_env.a);",
                "    loopEnd3 = _.NUM(_env.b);",
                "    loopConstraintsInitialised4 = true;",
                "});",
                "if (loopConstraintsInitialised4 && (loopStart2 <= loopEnd3))",
                "{",
                "    _.HANDLEERROR(errOn1, () =>",
                "    {",
                "        for (_env.i = loopStart2; _.NUM(_env.i) <= loopEnd3; _env.i = _.NUM(_env.i) + 1)",
                "        {",
                "            _.HANDLEERROR(errOn1, () => {",
                "                _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.i, v5 => { _env.i = v5; }));",
                "            });",
                "        }",
                "    });",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop constraints and known numeric values at translation time then enabling error-handling is relatively easy. The loop needs to be
        /// wrapped in error-trapping in case the termination condition or loop variable addition/subtraction fail, then the individual statements
        /// within the loop need wrapping as well. But without any dynamic loop constraints to be evaluated, it's a lot simpler - no evaluation
        /// of contraints to trap or guard clause around the loop to worry about.
        /// </summary>
        [Fact]
        public void AscendingLoopWithImplicitStepAndErrorTrappingEnabled()
        {
            var source = @"
                On Error Resume Next
                For i = 1 To 10
                    WScript.Echo i
                Next
            ";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPING(errOn1);",
                "_.HANDLEERROR(errOn1, () =>",
                "{",
                "    for (_env.i = 1; _.NUM(_env.i) <= 10; _env.i = _.NUM(_env.i) + 1)",
                "    {",
                "        _.HANDLEERROR(errOn1, () => {",
                "            _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.i, v2 => { _env.i = v2; }));",
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

        // TODO: Various variable-ascending/descending/step combinations
    }
}
