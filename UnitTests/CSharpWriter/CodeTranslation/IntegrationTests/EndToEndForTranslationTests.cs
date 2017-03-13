using System;
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
		/// VBScript, the constraints are not re-evaluated each loop iteration). The loop may only be entered if there is a zero or positive step and a
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
				"   loopEnd2 = _.NUM(_env.b);",
				"   loopStart3 = _.NUM(_env.a);",
				"   if ((loopStart3 is DateTime) || (loopStart3 is Decimal))",
				"       _env.i = loopStart3;",
				"   loopStart3 = _.NUM(_env.a, loopEnd2, (Int16)1);",
				"   loopConstraintsInitialised4 = true;",
				"});",
				"if (_.StrictLTE(loopStart3, loopEnd2))",
				"{",
				"   if (loopConstraintsInitialised4)",
				"       _env.i = loopStart3;",
				"   while (true)",
				"   {",
				"       _.HANDLEERROR(errOn1, () => {",
				"           _.CALL(this, _env.wscript, \"Echo\", _.ARGS.Ref(_env.i, v5 => { _env.i = v5; }));",
				"       });",
				"       if (!loopConstraintsInitialised4)",
				"           break;",
				"       var continueLoop6 = false;",
				"       _.HANDLEERROR(errOn1, () => {",
				"           _env.i = _.ADD(_env.i, (Int16)1);",
				"           continueLoop6 = _.StrictLTE(_env.i, loopEnd2);",
				"       });",
				"       if (!continueLoop6)",
				"           break;",
				"   }",
				"}",
				"_.RELEASEERRORTRAPPINGTOKEN(errOn1);"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If the loop constraints are known numeric values at translation time then enabling error-handling is relatively easy. The loop needs to be
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
				"_env.i = (Int16)1;",
				"while (true)",
				"{",
				"   _.HANDLEERROR(errOn1, () => {",
				"       _.CALL(this, _env.wscript, \"Echo\", _.ARGS.Ref(_env.i, v2 => { _env.i = v2; }));",
				"   });",
				"   var continueLoop3 = false;",
				"   _.HANDLEERROR(errOn1, () => {",
				"       _env.i = _.ADD(_env.i, (Int16)1);",
				"       continueLoop3 = _.StrictLTE(_env.i, 10);",
				"   });",
				"   if (!continueLoop3)",
				"       break;",
				"}",
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

		// TODO: Various variable-ascending/descending/step combinations

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
				"public object f1()",
				"{",
				"    object retVal1 = null;",
				"    object j = null; /* Undeclared in source */",
				"    object i = null; /* Undeclared in source */",
				"    for (i = (Int16)1; _.StrictLTE(i, 5); i = _.ADD(i, (Int16)1))",
				"    {",
				"        _.CALL(this, _env.wscript, \"Echo\", _.ARGS.Ref(j, v2 => { j = v2; }));",
				"    }",
				"    return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a FOR loop exists within a function F1 and one of F1's arguments is used to determine a loop constraint and that argument is passed to F1 ByRef and that argument is passed
		/// to another function while determining the loop constraints then a ByRef mapping will be required for the F1 argument. This is because the argument will be referenced inside a
		/// lambda when passed as Ref argument and it is not legal C# to reference a ref argument within a lambda.
		/// </summary>
		[Fact]
		public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired()
		{
			var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
				End Function";

			var expected = @"
				public object f1(ref object x)
				{
					object retVal1 = null;
					object i = null;
					object loopEnd4 = 0, loopStart5 = 0;
					var loopConstraintsInitialised6 = false;
					object byrefalias2 = x;
					try
					{
						loopEnd4 = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias2, v3 => { byrefalias2 = v3; })));
						loopStart5 = _.NUM((Int16)1);
						if ((loopStart5 is DateTime) || (loopStart5 is Decimal))
							i = loopStart5;
						loopStart5 = _.NUM((Int16)1, loopEnd4);
						loopConstraintsInitialised6 = true;
					}
					finally { x = byrefalias2; }
					if (_.StrictLTE(loopStart5, loopEnd4))
					{
						for (i = loopStart5; _.StrictLTE(i, loopEnd4); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return retVal1;
				}

				public object f2(ref object value)
				{
					return _.VAL(value);
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This test is a companion of IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired and illustrates that a ByRef mapping is not required
		/// for the F1 argument if the argument is passed in ByVal (while the argument still needs to be referenced in a lambda when passed to F2 as a ByRef argument, it's not a ref argument in
		/// F1 and so we don't need to jump through any hoops to avoid illegal C#)
		/// </summary>
		[Fact]
		public void IfByValArgumentIsRequiredForLoopConstraintAndIsPassedToAnotherFunctionByRefThenNoByRefMappingIsRequiredAsTheFirstArgumentWasByVal()
		{
			var source = @"
				Function F1(ByVal x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
				End Function";

			var expected = @"
				public object f1(object x)
				{
					object retVal1 = null;
					object i = null;
					var loopEnd3 = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(x, v2 => { x = v2; })));
					var loopStart4 = _.NUM((Int16)1, loopEnd3);
					if (_.StrictLTE(loopStart4, loopEnd3))
					{
						for (i = loopStart4; _.StrictLTE(i, loopEnd3); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return retVal1;
				}

				public object f2(ref object value)
				{
					return _.VAL(value);
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is a companion to IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired and illustrates a limitation of the translation process.
		/// The function F1 takes a ByRef argument which is then passed to F2 as the loop constraints are initialised. Although F2 accepts the argument ByVal, the translation analysis does not
		/// go deeply enough to realise this and presumes that F2 may take the argument ByRef - as such, it tries to pass it ByRef (just in case) and so needs to reference the F1 argument within
		/// a lambda, which would not be legal C# and so a ByRef mapping is unfortunately required.
		/// </summary>
		[Fact]
		public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionThenByRefMappingRequired()
		{
			var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByVal value)
					F2 = value
				End Function";

			var expected = @"
				public object f1(ref object x)
				{
					object retVal1 = null;
					object i = null;
					object loopEnd4 = 0, loopStart5 = 0;
					var loopConstraintsInitialised6 = false;
					object byrefalias2 = x;
					try
					{
						loopEnd4 = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias2, v3 => { byrefalias2 = v3; })));
						loopStart5 = _.NUM((Int16)1);
						if ((loopStart5 is DateTime) || (loopStart5 is Decimal))
							i = loopStart5;
						loopStart5 = _.NUM((Int16)1, loopEnd4);
						loopConstraintsInitialised6 = true;
					}
					finally { x = byrefalias2; }
					if (_.StrictLTE(loopStart5, loopEnd4))
					{
						for (i = loopStart5; _.StrictLTE(i, loopEnd4); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return retVal1;
				}

				public object f2(object value)
				{
					return _.VAL(value);
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is a companion to IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionThenByRefMappingRequired and shows that we can make things a little better by
		/// presuming that all built-in functions take arguments ByVal (which I'm fairly confident is always the case), which means that ByRef mappings may be avoided for some cases.
		/// </summary>
		[Fact]
		public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToBuiltInFunctionByRefThenNoByRefMappingRequired()
		{
			var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To UBOUND(x)
					Next
				End Function";

			var expected = @"
				public object f1(ref object x)
				{
					object retVal1 = null;
					object i = null;
					var loopEnd2 = _.UBOUND(x);
					var loopStart3 = _.NUM((Int16)1, loopEnd2);
					if (_.StrictLTE(loopStart3, loopEnd2))
					{
						for (i = loopStart3; _.StrictLTE(i, loopEnd2); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return retVal1;
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a ByRef argument is passed to a method F1 that uses the argument when evaluating loop constraints within an error-trapping block then a ByRef mapping will be required because the
		/// ByRef argument will need to be accessed within a lambda (inside the HANDLEERROR block), which is not legal in C#.
		/// </summary>
		[Fact]
		public void IfByRefArgumentIsRequiredForKnownLoopConstraintsAndLoopWrappedInErrorTrappingThenByRefMappingRequired()
		{
			var source = @"
				Function F1(ByRef x)
					On Error Resume Next
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
					value = 123
				End Function";

			var expected = @"
				public object f1(ref object x)
				{
					object retVal1 = null;
					var errOn2 = _.GETERRORTRAPPINGTOKEN();
					object i = null;
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);
					object loopEnd5 = 0, loopStart6 = 0;
					var loopConstraintsInitialised7 = false;
					object byrefalias3 = x;
					try
					{
						_.HANDLEERROR(errOn2, () => {
							loopEnd5 = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias3, v4 => { byrefalias3 = v4; })));
							loopStart6 = _.NUM((Int16)1);
							if ((loopStart6 is DateTime) || (loopStart6 is Decimal))
								i = loopStart6;
							loopStart6 = _.NUM((Int16)1, loopEnd5);
							loopConstraintsInitialised7 = true;
						});
					}
					finally { x = byrefalias3; }
					if (_.StrictLTE(loopStart6, loopEnd5))
					{
						if (loopConstraintsInitialised7)
							i = loopStart6;
						while (true)
						{
							if (!loopConstraintsInitialised7)
								break;
							var continueLoop8 = false;
							_.HANDLEERROR(errOn2, () => {
								i = _.ADD(i, (Int16)1);
								continueLoop8 = _.StrictLTE(i, loopEnd5);
							});
							if (!continueLoop8)
								break;
						}
					}
					_.RELEASEERRORTRAPPINGTOKEN(errOn2);
					return retVal1;
				}

				public object f2(ref object value)
				{
					object retVal9 = null;
					retVal9 = _.VAL(value);
					value = (Int16)123;
					return retVal9;
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a ByRef argument is passed to a method F1 that uses the argument when evaluating loop constraints within an error-trapping block then a ByRef mapping will be required because the
		/// ByRef argument will need to be accessed within a lambda (inside the HANDLEERROR block), which is not legal in C#. When it is known that the ByRef argument will not be changed by the
		/// loop constraint evaluation, the ByRef mapping is readonly; meaning that no try..finally wrapping is required to write the byref-temp-value back over the method argument (because it
		/// is known that the temporary value will not have been manipulated).
		/// </summary>
		[Fact]
		public void IfByRefArgumentIsRequiredForKnownReadOnlyLoopConstraintsAndLoopWrappedInErrorTrappingThenReadOnlyByRefMappingRequired()
		{
			var source = @"
				Function F1(ByRef x)
					On Error Resume Next
					Dim i: For i = 1 To x + 1
					Next
				End Function";

			var expected = @"
				public object f1(ref object x)
				{
					object retVal1 = null;
					var errOn2 = _.GETERRORTRAPPINGTOKEN();
					object i = null;
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);
					object loopEnd4 = 0, loopStart5 = 0;
					var loopConstraintsInitialised6 = false;
					object byrefalias3 = x;
					_.HANDLEERROR(errOn2, () => {
						loopEnd4 = _.NUM(_.ADD(byrefalias3, (Int16)1));
						loopStart5 = _.NUM((Int16)1);
						if ((loopStart5 is DateTime) || (loopStart5 is Decimal))
							i = loopStart5;
						loopStart5 = _.NUM((Int16)1, loopEnd4);
						loopConstraintsInitialised6 = true;
					});
					if (_.StrictLTE(loopStart5, loopEnd4))
					{
						if (loopConstraintsInitialised6)
							i = loopStart5;
						while (true)
						{
							if (!loopConstraintsInitialised6)
								break;
							var continueLoop7 = false;
							_.HANDLEERROR(errOn2, () => {
								i = _.ADD(i, (Int16)1);
								continueLoop7 = _.StrictLTE(i, loopEnd4);
							});
							if (!continueLoop7)
								break;
						}
					}
					_.RELEASEERRORTRAPPINGTOKEN(errOn2);
					return retVal1;
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}
	}
}
