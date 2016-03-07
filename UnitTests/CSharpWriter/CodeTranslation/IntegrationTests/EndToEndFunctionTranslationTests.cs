using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndFunctionTranslationTests
	{
		/// <summary>
		/// When a function (or property) has only a single executable statement that is a return-this-expression statement, then this can be translated
		/// into a single line C# return statement. Anything more complicated requires a temporary variable which is used to track the return value and
		/// returned from any exit point.
		/// </summary>
		[Fact]
		public void IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatement()
		{
			var source = @"
				PUBLIC FUNCTION F1()
					' Test simple-return-format functions
					F1 = CDate(""2007-04-01"")
				END FUNCTION
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"    // Test simple-return-format functions",
				"    return _.CDATE(\"2007-04-01\");",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is similar to IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatement but the single value passed a ByRef
		/// argument of the containing function into another function as a ByRef argument - as such, it will be referenced within a lambda. This is
		/// not acceptable in C# and so the argument must be stored in an alias while the second call takes place and then mapped back over the
		/// original value. As such, it can not be represented as a simple one-line return statement.
		/// </summary>
		[Fact]
		public void IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatementUnlessRefAliasMappingsRequired()
		{
			var source = @"
				PUBLIC FUNCTION F1(a)
					F1 = F2(a)
				END FUNCTION
				PUBLIC FUNCTION F2(a)
				END FUNCTION
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal2 = null;",
				"    object byrefalias4 = a;",
				"    try",
				"    {",
				"        retVal2 = _.VAL(_.CALL(_outer, \"F2\", _.ARGS.Ref(byrefalias4, v5 => { byrefalias4 = v5; })));",
				"    }",
				"    finally { a = byrefalias4; }",
				"    return retVal2;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is very similar to IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatement except that it demonstrates the difference
		/// required when the function return value is SET - meaning that it must be an object reference (however, because it references an undeclared variable,
		/// that variable must be defined within the function scope; so the C# is no longer a one-executable-line job, but the principle remains).
		/// </summary>
		[Fact]
		public void IfTheOnlyExecutableStatementIsSetReturnValueThenTranslateIntoSingleReturnStatement()
		{
			var source = @"
				PUBLIC FUNCTION F1()
					Set F1 = a
				END FUNCTION
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"    object a = null; /* Undeclared in source */",
				"    return _.OBJ(a);",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is very similar to IfTheOnlyExecutableStatementIsSetReturnValueThenTranslateIntoSingleReturnStatement except that it demonstrates the fact that
		/// if the return reference is already known to be an object type (which "Nothing" is) then it doesn't need to call the OBJ method for safety.
		/// </summary>
		[Fact]
		public void IfTheOnlyExecutableStatementIsSetKnownObjectReturnValueThenTranslateIntoSingleReturnStatement()
		{
			var source = @"
				PUBLIC FUNCTION F1()
					Set F1 = Nothing
				END FUNCTION
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"    return VBScriptConstants.Nothing;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a ByRef argument of a function is passed direct into another function as a ByRef argument then it must be stored in a temporary variable and then
		/// updated from this variable after the function call completes (whether it succeeds or fails - if the ByRef argument was altered before the error
		/// then that updated value must be persisted). This is to avoid trying to access "ref" reference in a lambda, which is a compile error in C#.
		/// Note that this applies only to ByRef arguments being passed onto another function as a ByRef argument - if a ByRef argument "a" is passed into
		/// another function indirectly (eg. with "a.Name") then it would not be passed to that second function ByRef and so the temporary "alias"
		/// variable would not be required.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsByRefArgument()
		{
			var source = @"
				Function F1(a)
					F2 a
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    object byrefalias2 = a;",
				"    try",
				"    {",
				"        _.CALL(_outer, \"F2\", _.ARGS.Ref(byrefalias2, v3 => { byrefalias2 = v3; }));",
				"    }",
				"    finally { a = byrefalias2; }",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This works as a counter-example to ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsByRefArgument, illustrating that if
		/// a ByRef argument of the containing function is only indirectly passed into a function call as an argument that would be ByRef then the alias
		/// variable is not required.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentNeedNotBeMappedToAnAliasIfIndirectlyAccessedOutsideOfErrorTrapping()
		{
			var source = @"
				Function F1(a)
					F2 a.Name
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    _.CALL(_outer, \"F2\", _.ARGS.Val(_.CALL(a, \"Name\")));",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This acts as a complement to the tests ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsByRefArgument and
		/// ByRefFunctionArgumentNeedNotBeMappedToAnAliasIfIndirectlyAccessedOutsideOfErrorTrapping by reminding us that a variable is not considered to
		/// be "indirectly" accessed if it looks like an array access - we can't know until runtime whether it is an array access or whether it was a
		/// default member access, meaning that it needs to be passed inside a REF-like lambda and so is another cases where an alias is required.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsAnArrayAsByRefArgument()
		{
			var source = @"
				Function F1(a)
					F2 a(0)
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    object byrefalias2 = a;",
				"    try",
				"    {",
				"        _.CALL(_outer, \"F2\", _.ARGS.RefIfArray(byrefalias2, _.ARGS.Val((Int16)0)));",
				"    }",
				"    finally { a = byrefalias2; }",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is only really here to go with ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsAnArrayAsByRefArgument and the test that
		/// that works with, it seemed like an obvious variation on the VBScript source - to follow the variable access with brackets even though there are no
		/// arguments (call behaviour around zero-argument brackets has just been changed so this is important). Here, the F1 argument will be accessed as an
		/// array (with no indexes specified) or a method - both of which will fail at runtime in VBScript. But the point is that the argument is not then
		/// passed on by-ref to F2 and so needs none of the special by-ref aliasing magic applying.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentRequiresNoSpecialTreatmentIfAccessAsFunctionOrArray()
		{
			var source = @"
				Function F1(a)
					F2 a()
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    _.CALL(_outer, \"F2\", _.ARGS.Val(_.CALL(a, _.ARGS.ForceBrackets())));",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This works as a another counter-example to ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsByRefArgument, if the ByRef
		/// argument of the containing function constitues only part of a value passed into a function call as an argument that would be ByRef then the alias
		/// variable is not required.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentNeedNotBeMappedToAnAliasIfOnlyPartialArgumentValue()
		{
			var source = @"
				Function F1(a)
					F2 """" & a
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    _.CALL(_outer, \"F2\", _.ARGS.Val(_.CONCAT(\"\", a)));",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This works as a another counter-example to ByRefFunctionArgumentRequiresSpecialTreatmentIfDirectlyUsedElsewhereAsByRefArgument, if the ByRef
		/// argument of the containing function is wrapped in extra brackets then it will be treated as ByVal even if otherwise it would need to be
		/// considered as ByRef when passed into another function.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentNeedNotBeMappedToAnAliasIfForcedInToByValWhenPassedToNextFunction()
		{
			var source = @"
				Function F1(a)
					F2 (a)
				End Function

				Function F2(a)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    _.CALL(_outer, \"F2\", _.ARGS.Val(a));",
				"    return retVal1;",
				"}",
				"public object f2(ref object a)",
				"{",
				"    return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a ByRef argument of a function is required within an expression that must potentially trap errors, then an alias of that argument will be
		/// required since the potentially-error-trapped content will be executed within a lambda and C# does not allow "ref" arguments to be accessed
		/// within lambdas. If this alias may be altered - if it is passed into another function as a ByRef argument, for example - then the alias value
		/// must be used to overwrite the original function argument reference, even if the expression evaluation failed (since it might have changed
		/// the value before the error occurred).
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentMustBeMappedToReadAndWriteAliasIfReferencedInReadAndWriteMannerWithinPotentiallyErrorTrappingStatement()
		{
			var source = @"
				Function F1(a)
					On Error Resume Next
					WScript.Echo a
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    var errOn2 = _.GETERRORTRAPPINGTOKEN();",
				"    _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);",
				"    object byrefalias3 = a;",
				"    try",
				"    {",
				"        _.HANDLEERROR(errOn2, () => {",
				"            _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(byrefalias3, v4 => { byrefalias3 = v4; }));",
				"        });",
				"    }",
				"    finally { a = byrefalias3; }",
				"    _.RELEASEERRORTRAPPINGTOKEN(errOn2);",
				"    return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a ByRef argument of a function is required within an expression that must potentially trap errors, then an alias of that argument will be
		/// required since the potentially-error-trapped content will be executed within a lambda and C# does not allow "ref" arguments to be accessed
		/// within lambdas. If this alias may not be altered then the alias need not be written back over the original reference, it is a read-only
		/// alias. This would be the case if there is a ByRef argument "a" of the current function and "a.Name" is passed to another function (as a
		/// ByRef OR ByVal argument) since the "a" in "a.Name" can never be affected.
		/// </summary>
		[Fact]
		public void ByRefFunctionArgumentMustBeMappedToReadOnlyAliasIfReferencedInReadOnlyMannerWithinPotentiallyErrorTrappingStatement()
		{
			var source = @"
				Function F1(a)
					On Error Resume Next
					WScript.Echo a.Name
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object a)",
				"{",
				"    object retVal1 = null;",
				"    var errOn2 = _.GETERRORTRAPPINGTOKEN();",
				"    _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);",
				"    object byrefalias3 = a;",
				"    _.HANDLEERROR(errOn2, () => {",
				"        _.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CALL(byrefalias3, \"Name\")));",
				"    });",
				"    _.RELEASEERRORTRAPPINGTOKEN(errOn2);",
				"    return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// Within a function F1, if there are places where it looks like F1 is being used as a variable (such as passing it as an argument into
		/// another function) then the C# code must use the temporary-return-value that is generated for the function (in the example here, the
		/// code IsEmpty(F1) should check the current return value of F1 for Empty, it shouldn't try to call F1 as a function and pass the
		/// result into IsEmpty)
		/// </summary>
		[Fact]
		public void ParentFunctionNameMustBeMappedOnToReturnValueNameWhenPassedAsArgumentToOtherFunction()
		{
			var source = @"
				Function F1(x)
					If IsEmpty(F1) Then
						F1 = True
					End If
				End Function

				Function F2(x)
				End Function
			";
			var expected = new[]
			{
				"public object f1(ref object x)",
				"{",
				"	object retVal1 = null;",
				"	if (_.IF(_.ISEMPTY(retVal1)))",
				"	{",
				"		retVal1 = true;",
				"	}",
				"	return retVal1;",
				"}",
				"public object f2(ref object x)",
				"{",
				"	return null;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If what looks like a function-return-value-setting statement has brackets after the function name then it's a type mismatch (and even
		/// even if it's only a single line, it won't be considered a simple short cut case since the analysis to determine whether that is the
		/// case or not is very basic; if the target is not a single NameToken that corresponds to the function name then the short cut is
		/// not applied)
		/// </summary>
		[Fact]
		public void IncludingBracketsWhenSettingTheReturnValueIsTypeMismatch()
		{
			var source = @"
				PUBLIC FUNCTION F1()
					F1() = Null
				END FUNCTION
			";
			var expected = new[]
			{
				"public object f1()",
				"{",
				"    object retVal1 = null;",
				"    _.SET(VBScriptConstants.Null, _.RAISEERROR(new TypeMismatchException(\"'F1'\")));",
				"    return retVal1;",
				"}"
			};
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}
	}
}
