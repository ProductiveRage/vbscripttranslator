using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndMiscTranslationTests
    {
        // TODO: Test function call with numeric values (1 and 1.1), string values, built-in values and built-in functions (such as "Now") and ensure that
        // they all have the arguments specified as ByVal
        // - Is it easiest to put it in here or better to put it into StatementTranslatorTests?

        /// <summary>
        /// The code here accesses an undeclared variable in a statement in the outermost scope, that scope should be registered in the EnvironmentReferences
        /// class. There is also a "wscript" reference which is declared as an External Dependency in the translator, this will appear in the Environment
        /// References class as well (as any/all External Dependencies should).
        /// </summary>
        [Fact]
        public void UndeclaredVariablesInTheOutermostScopeShouldBeDefinedAsAnEnvironmentVariable()
        {
            var source = @"
                WScript.Echo i
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(_env.i, v1 => { _env.i = v1; }));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This code will access an undeclared variable within a function. The scope of that undeclared variable should be restricted to the function in
        /// which it is accessed and not bleed out into the outer scope.
        /// </summary>
        [Fact]
        public void UndeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"Test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null; /* Undeclared in source */",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    Dim i
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"Test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null;",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableInOutermostScopeShouldBeAccessedFromThereWhenRequiredWithinFunction()
        {
            var source = @"
                Dim i
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"Test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(_outer.i, v2 => { _outer.i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void NumericLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func 1()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"'[number: 1]' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void StringLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func \"1\"()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"'[string: \\\"1\\\"]' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void BuiltinValuesAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func vbObjectError()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"'vbObjectError' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void ClassNameFollowedByBracketsInNewStatementResultsInCompileTimeError()
        {
            var source = "c = new C1()";
            Assert.Throws<Exception>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        /// <summary>
        /// Since runs of string concatenations are so common, an exception to the two-arguments-per-operation (apart from NOT, that only takes one) is made
        /// to allow the values to be combined in a single CONCAT call, reducing the size of the emitted code
        /// </summary>
        [Fact]
        public void ConcatFunctionAllowsMoreThanTwoArguments()
        {
            var source = @"
                WScript.Echo a & b & c & d
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CONCAT(_env.a, _env.b, _env.c, _env.d)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is related to the ConcatFunctionAllowsMoreThanTwoArguments and provides reassurance that string concatenations will only be joined if it
        /// would have no effect on the rest of processing (since the addition operation should take precedence, there is no CONCAT-flattening that can
        /// be performed in this case)
        /// </summary>
        [Fact]
        public void ConcatFunctionAllowsMoreThanTwoArgumentsButDoesNotAffectNestedOperationsOfOtherTypes()
        {
            var source = @"
                WScript.Echo a & 1 + 2 & c & d
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CONCAT(_env.a, _.ADD((Int16)1, (Int16)2), _env.c, _env.d)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// The string values that specify target member names in a CALL expression must not be manipulated by the name rewriter at runtime. This means that
        /// their casing will not be affected and - more importantly, any manipulations relating to C# keywords will NOT be applied. When the target is a
        /// translated class, the name rewriter manipulations would not cause any issue but if the target is not something that is translated (a COM component,
        /// for example), then trying to access its members with the name-rewritten versions will fail. This means that the CALL implementation must be able to
        /// consider the same name rewriter rules at runtime that the translator does.
        /// </summary>
        [Fact]
        public void MemberAccessorsInCallStatementsShouldNotBeRenamedAtTranslationTime()
        {
            // "Params" is a C# keyword, so we couldn't emit translated code with a method called "Params", but if "a" is an external reference (such as a COM
            // component) then It may have a methor or property named "Params". As such we mustn't enforce the rewriting of "Params" to something C#-friendly
            // at compile time (the CALL implementation will have to do some magic)
            // - The GetTranslatedStatements uses the DefaultTranslator which uses the DefaultRuntimeSupportClassFactory.DefaultNameRewriter which will
            //   ensure that C# keywords are rewritten to something safe
            var source = @"
                WScript.Echo a.Params
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CALL(_env.a, \"Params\")));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// Similar to MemberAccessorsInCallStatementsShouldNotBeRenamedAtTranslationTime, the ValueSettingStatementsTranslator has been corrected so that it
        /// won't rewrite member accessors that string arguments in a SET call
        /// </summary>
        [Fact]
        public void MemberAccessorsInValueSettingsStatementsShouldNotBeRenamedAtTranslationTime()
        {
            var source = @"
                a.Name.Length = 1
            ";
            var expected = new[]
            {
                "_.SET((Int16)1, _.CALL(_env.a, \"Name\"), \"Length\");"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// It doesn't matter if we're within a VBScript class on in the outermost scope, or within a function in the outermost scope, the "Me" reference may
        /// always be mapped directly to "this" and it will be correct
        /// </summary>
        [Fact]
        public void MeReferenceMapsDirectlyOnToThis()
        {
            var source = @"
                WScript.Echo Me.Name
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CALL(this, \"Name\")));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a CALL expression has a function as its target then it needs to be rewritten so that that the owner of the function (or property) is the target
        /// and the function and one of the member accessors (since it's not valid C# to provide a delegate for an object argument)
        /// </summary>
        [Fact]
        public void OutermostScopeFunctionMayNotBeTargetOfCallExpression()
        {
            var source = @"
                Set a = GetSomething.Name
                Function GetSomething()
                End Function
            ";
            var expected = new[]
            {
                "_env.a = _.OBJ(_.CALL(_outer, \"GetSomething\", \"Name\"));",
                "public object getsomething()",
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
        /// Very similar to OutermostScopeFunctionMayNotBeTargetOfCallExpression except that the function is within a class rather than in the outermost scope
        /// </summary>
        [Fact]
        public void ClassContainedFunctionMayNotBeTargetOfCallExpression()
        {
            var source = @"
                Class C1
                    Function Go()
                        Set a = GetSomething.Name
                    End Function
                    Function GetSomething()
                    End Function
                End Class";
            var expected = @"
                [ComVisible(true)]
                [SourceClassName(""C1"")]
                public sealed class c1
                {
                    private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
                    private readonly EnvironmentReferences _env;
                    private readonly GlobalReferences _outer;
                    public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
                    {
                        if (compatLayer == null)
                            throw new ArgumentNullException(""compatLayer"");
                        if (env == null)
                            throw new ArgumentNullException(""env"");
                        if (outer == null)
                            throw new ArgumentNullException(""outer"");
                        _ = compatLayer;
                        _env = env;
                        _outer = outer;
                    }

                    public object go()
                    {
                        object retVal1 = null;
                        object a = null; /* Undeclared in source */
                        a = _.OBJ(_.CALL(this, ""GetSomething"", ""Name""));
                        return retVal1;
                    }
                    public object getsomething()
                    {
                        return null;
                    }
                }";
            Assert.Equal(
                expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

		/// <summary>
		/// This illustrated a bug that was identified with numeric literals of the form "&H001" - the trailing zeroes were causing an exception in the
		/// parsing process
		/// </summary>
		[Fact]
		public void EnsureThatHexValuesWithTrailingZeroesAreParsedCorrectly()
		{
			var source = @"
                const SOME_CONSTANT = &H0001
            ";
			var expected = new[]
            {
                "_outer.some_constant = 1;"
            };
			Assert.Equal(
				expected.Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

        [Theory, MemberData("VariousBracketDeterminedRefValArgumentData")]
        public void VariousBracketDeterminedRefValArgumentCases(string source, string expectedResult)
        {
            var translatedContent = WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            Assert.Equal(expectedResult, translatedContent.Select(c => c.Trim()).Single(c => c != ""));
        }

        public static IEnumerable<object[]> VariousBracketDeterminedRefValArgumentData
        {
            get
            {
                yield return new object[] { "func x", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }));" };
                yield return new object[] { "func (x)", "_.CALL(_env.func, _.ARGS.Val(_env.x));" };

                yield return new object[] { "func x, y", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Ref(_env.y, v2 => { _env.y = v2; }));" };
                yield return new object[] { "func (x), y", "_.CALL(_env.func, _.ARGS.Val(_env.x).Ref(_env.y, v1 => { _env.y = v1; }));" };
                yield return new object[] { "func x, (y)", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Val(_env.y));" };

                yield return new object[] { "z = func(x)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; })));" };
                yield return new object[] { "z = func(x, y)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Ref(_env.y, v2 => { _env.y = v2; })));" };
                yield return new object[] { "z = func((x), y)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Val(_env.x).Ref(_env.y, v1 => { _env.y = v1; })));" };
            }
        }

        [Theory, MemberData("ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessaryData")]
        public void ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessary(string source, string expectedResult)
        {
            var translatedContent = WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            Assert.Equal(expectedResult, translatedContent.Select(c => c.Trim()).Single(c => c != ""));
        }

        public static IEnumerable<object[]> ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessaryData
        {
            get
            {
                yield return new object[] { "a = b", "_env.a = _.VAL(_env.b);" };
                yield return new object[] { "a = b()", "_env.a = _.VAL(_.CALL(_env.b, _.ARGS.ForceBrackets()));" };
                yield return new object[] { "a = b(1)", "_env.a = _.VAL(_.CALL(_env.b, _.ARGS.Val((Int16)1)));" };

                yield return new object[] { "a = b.Name", "_env.a = _.VAL(_.CALL(_env.b, \"Name\"));" };
                yield return new object[] { "a = b.Name()", "_env.a = _.VAL(_.CALL(_env.b, \"Name\", _.ARGS.ForceBrackets()));" };
                yield return new object[] { "a = b.Name(1)", "_env.a = _.VAL(_.CALL(_env.b, \"Name\", _.ARGS.Val((Int16)1)));" };
            }
        }
    }
}
