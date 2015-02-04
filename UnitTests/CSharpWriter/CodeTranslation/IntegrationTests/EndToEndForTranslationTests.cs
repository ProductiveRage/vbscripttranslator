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
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A loop that exceeds the range of the VBScript "Integer" will result in the loop variable being set to a larger type so that it can describe all of the
        /// values within the loop. Note that there is special handling to identify the case when all loop constraints are constants within VBScript's "Integer"
        /// range, which is why the test AscendingLoopWithImplicitStep does not require an addition "loopStart" variable. That shortcut is not in play here
        /// and so a "loopStart" variable IS required (to determine what type to use to cover the range from (Int16)1 to 32768 - which is implicitly an
        /// Int32 (aka "int") when compiled as C#.
        /// </summary>
        [Fact]
        public void AscendingLoopThatRollsOverLoopVariableIntoLongType()
        {
            var source = @"
                Dim i: For i = 1 To 32768
                Next
            ";
            var expected = new[]
            {
                "var loopStart1 = _.NUM((Int16)1, 32768);",
                "for (_outer.i = loopStart1; _.StrictLTE(_outer.i, 32768); _outer.i = _.ADD(_outer.i, (Int16)1))",
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
                "for (_outer.i = (Int16)5; _.StrictGTE(_outer.i, 1); _outer.i = _.SUBT(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            Assert.Equal(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A fractional step on an otherwise small integer range changes the loop variable from being a VBScript "Integer" to a "Double"
        /// </summary>
        [Fact]
        public void DescendingLoopWithExplicitFractionalStep()
        {
            var source = @"
                Dim i: For i = 1 To 5 Step 0.1
                Next
            ";
            var expected = new[]
            {
                "var loopStart1 = _.NUM((Int16)1, (Int16)5, 0.1);",
                "for (_outer.i = loopStart1; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, 0.1))",
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
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5);)",
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
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, (Int16)2))",
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
                "for (_outer.i = (Int16)5; _.StrictGTE(_outer.i, 1); _outer.i = _.SUBT(_outer.i, (Int16)1))",
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
                "var loopEnd1 = _.NUM(_env.b);",
                "var loopStep2 = _.NUM(_env.c);",
                "var loopStart3 = _.NUM(_env.a, loopEnd1, loopStep2);",
                "if ((_.StrictLTE(loopStart3, loopEnd1) && _.StrictGTE(loopStep2, 0))",
                "|| (_.StrictGT(loopStart3, loopEnd1) && _.StrictLT(loopStep2, 0)))",
                "{",
                "for (_env.i = loopStart3;",
                "    (_.StrictGTE(loopStep2, 0) && _.StrictLTE(_env.i, loopEnd1)) || (_.StrictLT(loopStep2, 0) && _.StrictGTE(_env.i, loopEnd1));",
                "     _env.i = _.ADD(_env.i, loopStep2))",
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
        /// first. If these constraints are successfully evaluated, then the loop proceeds as would be expected, but there must be error-trapping around
        /// the loop itself, so that if the termination condition or loop-variable-addition/subtraction fails then the loop will terminate. There must
        /// also be error-trapping around each statement within the loop, so that if any one of them fails then the others may still be processed (if
        /// the error-trapping token is enabled at that point during the runtime execution). HOWEVER, the craziest bit is that if evaluation of any
        /// of the loop constraints fails then no further constraint evaluation will be attempted (they are processed in the order of From, To and
        /// Step) but the loop WILL be executed once. For this iteration, the loop variable will not be altered (so it will be left as null in
        /// this example, but if it had been set to "a" before the loop then it would remain set to "a") and neither the loop termination
        /// condition nor the increment work will be attempted.
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
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "object loopEnd2 = 0, loopStart3 = 0;",
                "var loopConstraintsInitialised4 = false;",
                "_.HANDLEERROR(errOn1, () => {",
                "    loopEnd2 = _.NUM(_env.b);",
                "    loopStart3 = _.NUM(_env.a);",
                "    if ((loopStart3 is DateTime) || (loopStart3 is Decimal))",
                "        _env.i = loopStart3;",
                "    loopStart3 = _.NUM(_env.a, loopEnd2, (Int16)1);",
                "    loopConstraintsInitialised4 = true;",
                "});",
                "if (!loopConstraintsInitialised4 || (_.StrictLTE(loopStart3, loopEnd2)))",
                "{",
                "    _.HANDLEERROR(errOn1, () => {",
                "        for (_env.i = loopConstraintsInitialised4 ? loopStart3 : _env.i;",
                "             !loopConstraintsInitialised4 || (_.StrictLTE(_env.i, loopEnd2));",
                "             _env.i = _.ADD(_env.i, (Int16)1))",
                "        {",
                "            _.HANDLEERROR(errOn1, () => {",
                "                _.CALL(_env.WScript, \"Echo\", _.ARGS.Ref(_env.i, v5 => { _env.i = v5; }));",
                "            });",
                "            if (!loopConstraintsInitialised4)",
                "                break;",
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
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "_.HANDLEERROR(errOn1, () => {",
                "    for (_env.i = (Int16)1; _.StrictLTE(_env.i, 10); _env.i = _.ADD(_env.i, (Int16)1))",
                "    {",
                "        _.HANDLEERROR(errOn1, () => {",
                "            _.CALL(_env.WScript, \"Echo\", _.ARGS.Ref(_env.i, v2 => { _env.i = v2; }));",
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

        /// <summary>
        /// A loop variable may be of type "Byte" but only if the start, end and step values are all of type "Byte" - if there is no step explicitly
        /// specified then the default "Integer" 1 will be used and so the loop variable will become type "Integer" (this test doesn't really show
        /// this completely since the translated code is not executed and it would depend upon the support class implementation but it seemed like
        /// it was worth recording here to make the point, also see the NUM test "BytesWithAnInteger")
        /// </summary>
        [Fact]
        public void ByteLoopStartAndEndValuesWithImplicitStepWillGetAnIntegerStep()
        {
            var source = @"
                Dim i: For i = CByte(1) To CByte(5)
                Next
            ";
            var expected = new[]
            {
                "var loopEnd1 = _.CBYTE(5);",
                "var loopStart2 = _.NUM(_.CBYTE(1), loopEnd1, (Int16)1);",
                "if (_.StrictLTE(loopStart2, loopEnd1))",
                "{",
                "    for (_outer.i = loopStart2; _.StrictLTE(_outer.i, loopEnd1); _outer.i = _.ADD(_outer.i, (Int16)1))",
                "    {",
                "    }",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is the complement to ByteLoopStartAndEndValuesWithImplicitStepWillGetAnIntegerStep, it illustrates how a loop would be constructed
        /// in order to have the loop variable be of type "Byte".
        /// </summary>
        [Fact]
        public void ByteLoopStartAndEndAndStepValuesWillGetByteLoopVariable()
        {
            var source = @"
                Dim i: For i = CByte(1) To CByte(5) Step CByte(1)
                Next
            ";
            var expected = new[]
            {
                "var loopEnd1 = _.CBYTE(5);",
                "var loopStep2 = _.CBYTE(1);",
                "var loopStart3 = _.NUM(_.CBYTE(1), loopEnd1, loopStep2);",
                "if ((_.StrictLTE(loopStart3, loopEnd1) && _.StrictGTE(loopStep2, 0))",
                "|| (_.StrictGT(loopStart3, loopEnd1) && _.StrictLT(loopStep2, 0)))",
                "{",
                "    for (_outer.i = loopStart3;",
                "        (_.StrictGTE(loopStep2, 0) && _.StrictLTE(_outer.i, loopEnd1)) || (_.StrictLT(loopStep2, 0) && _.StrictGTE(_outer.i, loopEnd1));",
                "         _outer.i = _.ADD(_outer.i, loopStep2))",
                "    {",
                "    }",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// When the translation of a for loop is completed, any undeclared variables should not be flushed (declared) within the scope of the loop, it
        /// must be within the scope-defining parent. If within a function then these must be local variables. This test also covers a fix where the loop
        /// variable was not getting identified as an undeclared variable when it should have been.
        /// </summary>
        [Fact]
        public void UndeclaredVariablesShouldNotBeFlushedAtForBlockEnd()
        {
            var source = @"
                Function F1
                    For i = 1 To 5
                        WScript.Echo j
                    Next
                End Function
            ";
            var expected = new[]
            {
                "public object F1()",
                "{",
                "    object retVal1 = null;",
                "    object j = null; /* Undeclared in source */",
                "    object i = null; /* Undeclared in source */",
                "    for (i = (Int16)1; _.StrictLTE(i, 5); i = _.ADD(i, (Int16)1))",
                "    {",
                "        _.CALL(_env.WScript, \"Echo\", _.ARGS.Ref(j, v2 => { j = v2; }));",
                "    }",
                "    return retVal1;",
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
